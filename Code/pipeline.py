from __future__ import annotations
import ast
import json
import re
from datetime import date
from typing import Any, cast
from itertools import combinations, chain
import httpx
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openai import OpenAI

def big_nba_boy(user_statement: str, api_key: str):
    """Filter NBA trade candidates, enumerate two-team exchanges, and apply user trade logic."""
    
    # STAGE 1 — NBA ROSTERS AND SPOTRAC TRADE RESTRICTIONS ONLY
    
    
    SPOTRAC_TEAMS = {
        "atlanta-hawks": ("Atlanta Hawks", "ATL"), "boston-celtics": ("Boston Celtics", "BOS"),
        "brooklyn-nets": ("Brooklyn Nets", "BKN"), "charlotte-hornets": ("Charlotte Hornets", "CHA"),
        "chicago-bulls": ("Chicago Bulls", "CHI"), "cleveland-cavaliers": ("Cleveland Cavaliers", "CLE"),
        "dallas-mavericks": ("Dallas Mavericks", "DAL"), "denver-nuggets": ("Denver Nuggets", "DEN"),
        "detroit-pistons": ("Detroit Pistons", "DET"), "golden-state-warriors": ("Golden State Warriors", "GSW"),
        "houston-rockets": ("Houston Rockets", "HOU"), "indiana-pacers": ("Indiana Pacers", "IND"),
        "la-clippers": ("Los Angeles Clippers", "LAC"), "los-angeles-lakers": ("Los Angeles Lakers", "LAL"),
        "memphis-grizzlies": ("Memphis Grizzlies", "MEM"), "miami-heat": ("Miami Heat", "MIA"),
        "milwaukee-bucks": ("Milwaukee Bucks", "MIL"), "minnesota-timberwolves": ("Minnesota Timberwolves", "MIN"),
        "new-orleans-pelicans": ("New Orleans Pelicans", "NOP"), "new-york-knicks": ("New York Knicks", "NYK"),
        "oklahoma-city-thunder": ("Oklahoma City Thunder", "OKC"), "orlando-magic": ("Orlando Magic", "ORL"),
        "philadelphia-76ers": ("Philadelphia 76ers", "PHI"), "phoenix-suns": ("Phoenix Suns", "PHX"),
        "portland-trail-blazers": ("Portland Trail Blazers", "POR"), "sacramento-kings": ("Sacramento Kings", "SAC"),
        "san-antonio-spurs": ("San Antonio Spurs", "SAS"), "toronto-raptors": ("Toronto Raptors", "TOR"),
        "utah-jazz": ("Utah Jazz", "UTA"), "washington-wizards": ("Washington Wizards", "WAS")
    }

    def normalize_name(value):
        value = re.sub(r"\s+", " ", str(value)).strip()
        value = re.sub(r"\s+[†‡*]+$", "", value).strip()
        value = re.sub(r"\s*\([^)]{1,8}\)\s*$", "", value).strip()
        return value

    TRADE_DATE = date.today()
    RESTRICTIONS_URL = "https://www.spotrac.com/nba/transactions/trade-restrictions"

    def collect_trade_restrictions():
        response = requests.get(
            RESTRICTIONS_URL,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/605.1.15 Version/18.6 Safari/605.1.15"
                ),
                "Accept": "text/html,application/xhtml+xml",
            },
            timeout=30,
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        team_aliases = {v[0].lower(): v[0] for v in SPOTRAC_TEAMS.values()}
        team_aliases["la clippers"] = "Los Angeles Clippers"
        rows = []
        title = soup.find(lambda t: t.name in {"h1", "h2"} and "nba trade restrictions" in t.get_text(" ", strip=True).lower())
        if title is None:
            print("WARNING: Spotrac restrictions heading unavailable; restriction list is empty")
            return pd.DataFrame(columns=["team", "Name", "restriction_date", "restriction_type"])
        for heading in title.find_all_next(["h2", "h3", "table"]):
            if heading.name in {"h2", "h3"} and heading.get_text(" ", strip=True).lower() in {"related", "more", "footer"}:
                break
            if heading.name != "table":
                continue
            
            previous = heading.find_previous(["h2", "h3", "h4"])
            raw_team = previous.get_text(" ", strip=True).strip() if previous is not None else ""
            team = team_aliases.get(raw_team.lower())
            if team is None:
                continue
            for tr in heading.find_all("tr"):
                cells = tr.find_all("td")
                if len(cells) < 3:
                    continue
                name = normalize_name(cells[0].get_text(" ", strip=True))
                
                player_link = cells[0].find("a")
                if player_link is not None:
                    name = normalize_name(player_link.get_text(" ", strip=True))
                expires = pd.to_datetime(cells[1].get_text(" ", strip=True), errors="coerce")
                reason = cells[2].get_text(" ", strip=True)
                if not name or pd.isna(expires) or not reason:
                    continue
                rows.append({"team": team, "Name": name, "restriction_date": expires.date(), "restriction_type": reason})
        result = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
        print(f"Spotrac restrictions: {len(result)} rows across {result["team"].nunique()} teams")
        return result

    
    # STAGE 1 — SCRAPE CURRENT NBA ROSTERS
    
    
    NBA_TEAM_IDS = {"ATL": 1610612737, "BOS": 1610612738, "BKN": 1610612751, "CHA": 1610612766, "CHI": 1610612741, "CLE": 1610612739, "DAL": 1610612742, "DEN": 1610612743, "DET": 1610612765, "GSW": 1610612744, "HOU": 1610612745, "IND": 1610612754, "LAC": 1610612746, "LAL": 1610612747, "MEM": 1610612763, "MIA": 1610612748, "MIL": 1610612749, "MIN": 1610612750, "NOP": 1610612740, "NYK": 1610612752, "OKC": 1610612760, "ORL": 1610612753, "PHI": 1610612755, "PHX": 1610612756, "POR": 1610612757, "SAC": 1610612758, "SAS": 1610612759, "TOR": 1610612761, "UTA": 1610612762, "WAS": 1610612764}

    def collect_nba_rosters():
        rosters = {}
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "text/html"})
        for _, (team, abbr) in SPOTRAC_TEAMS.items():
            url = f"https://www.nba.com/team/{NBA_TEAM_IDS[abbr]}"
            response = session.get(url, timeout=30)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            roster_table = next((table for table in soup.find_all("table") if {"player", "pos"}.issubset({cell.get_text(" ", strip=True).lower() for cell in table.find("tr").find_all(["th", "td"])}) if table.find("tr") is not None), None)
            if roster_table is None:
                print(f"WARNING: roster table missing for {team}")
                continue
            names = []
            for tr in roster_table.find_all("tr"):
                cells = tr.find_all("td")
                if not cells:
                    continue
                link = cells[0].find("a", href=re.compile(r"/player/|/players/"))
                name = normalize_name(link.get_text(" ", strip=True) if link else cells[0].get_text(" ", strip=True))
                if name and name.lower() != "player":
                    names.append(name)
            rosters[team] = sorted(set(names))
            print(f"NBA roster: {team} — {len(rosters[team])} players")
        return rosters

    league_players = collect_nba_rosters()
    trade_restrictions_df = collect_trade_restrictions()
    restriction_lookup = {}
    for _, row in trade_restrictions_df.iterrows():
        key = (str(row["team"]), normalize_name(row["Name"]).casefold())
        restriction_lookup.setdefault(key, []).append(row)

    
    # STAGE 2 — GPT PLAYER EVALUATIONS
    
    
    client = OpenAI(api_key=api_key, http_client=cast(Any, httpx.Client()))
    print(f"Scraped NBA rosters: {sum(map(len, league_players.values()))} players across {len(league_players)} teams")

    BATCH_SIZE = 5

    teams = list(league_players.items())

    all_results = []

    total_batches = (len(teams) + BATCH_SIZE - 1) // BATCH_SIZE


    for start_index in range(0, len(teams), BATCH_SIZE):

        batch_items = teams[start_index:start_index + BATCH_SIZE]

        batch = dict(batch_items)

        batch_text = json.dumps(batch, ensure_ascii=False, indent=2)

        batch_number = (start_index // BATCH_SIZE) + 1

        print(f"Player evaluation batch {batch_number}/{total_batches}")

        response = client.responses.create(
            model="gpt-5.6",
            tools=[{"type": "web_search"}],
            input=f"""
    Below is a list of current NBA players grouped by team.

    {batch_text}

    Evaluate EVERY player listed.

    For each player, return:

    1. Attributes:
       - Up to 8 concise attributes describing the player physically,
         mentally, stylistically, or by career stage.
       - Examples include:
         age,
         young,
         veteran,
         speed,
         quickness,
         strength,
         size,
         vertical athleticism,
         durability,
         temperament,
         basketball IQ,
         playoff experience,
         playoff riser,
         clutch,
         inconsistent,
         choker,
         high-motor,
         injury-prone,
         etc.

    2. Strengths:
       - Up to 5 meaningful CURRENT basketball strengths.
       - Possible areas include:
         shooting,
         three-point shooting,
         midrange shooting,
         finishing,
         rim finishing,
         passing,
         playmaking,
         ball handling,
         shot creation,
         off-ball movement,
         screening,
         rebounding,
         offensive rebounding,
         perimeter defense,
         interior defense,
         rim protection,
         switchability,
         athleticism,
         basketball IQ,
         decision making,
         transition offense,
         transition defense,
         post offense,
         post defense,
         etc.

    3. Weaknesses:
       - Up to 5 meaningful CURRENT basketball weaknesses.
       - Use the same types of basketball categories as the strengths.

    IMPORTANT RULES:

    - Be stringent.
    - Evaluate CURRENT ability only.
    - Do NOT evaluate players based on what they were at their historical peak.
    - Do NOT evaluate them based on what they might become in the future.
    - Do NOT manufacture strengths or weaknesses simply to reach 5.
    - If a player only has 2 clear weaknesses, return 2.
    - If a player only has 3 clear strengths, return 3.
    - Use current information when necessary.
    - Distinguish between reputation and current performance.
    - Do not call someone clutch or a choker unless there is meaningful evidence.
    - Do not infer personality traits without reasonable public evidence.
    - Basketball skills should be concrete enough that another program can
      later use them to evaluate trade preferences.
    - Evaluate every player in the supplied batch.
    - Do not omit players.

    Return the information EXACTLY in this general format:

    Team Name:

    Player Name: Attributes ["attribute 1", "attribute 2", ...], Strengths ["strength 1", "strength 2", ...], Weaknesses ["weakness 1", "weakness 2", ...]
    Player Name 2: Attributes ["attribute 1", "attribute 2", ...], Strengths ["strength 1", "strength 2", ...], Weaknesses ["weakness 1", "weakness 2", ...]

    Next Team Name:

    Player Name: Attributes [...], Strengths [...], Weaknesses [...]
    Player Name 2: Attributes [...], Strengths [...], Weaknesses [...]

    Return ONLY the requested team/player information.

    Do not include:
    - an introduction
    - an explanation
    - methodology
    - citations
    - source lists
    - conclusions
    - disclaimers
    """
        )

        batch_result = response.output_text.strip()

        all_results.append(batch_result)

    
    # STAGE 3 — PARSE THE USER'S TRADE PREFERENCES USING GPT
    
    
    demands = client.responses.create(
        model="gpt-5.6",
        input=f"""
You are interpreting an NBA trade request for a later player-pool filter.
Read the user statement and extract the following, without inventing names:

Team Name: [user's NBA team]
Required Acquisition Players: [specifically named players the user insists on acquiring]
Preferred Acquisition Players: [named players the user is interested in but does not require]
Acquisition Player Types: [roles, skills, traits of other acceptable acquisitions]
Preferred Outgoing Players: [named players the user explicitly wants to trade away or offers to trade]
Other Allowed Outgoing Players: [named players the user explicitly permits trading, without preferring to move]
Untouchable Players: [named players never to trade]
Untouchable Player Types: [traits of players never to trade]
Players to Avoid Acquiring: [specific names]
Player Types to Avoid Acquiring: [traits, weaknesses, roles to exclude]
Teams to Avoid: [teams the user refuses to deal with]
Required Trade Logic: [Python Boolean membership expressions, as quoted strings]

Extract mandatory logical conditions ONLY when explicitly stated by the user.
Use Python expressions with membership tests, AND/OR/NOT operators, parentheses,
and the variables acquired and outgoing. Use full player names, for example:
Required Trade Logic: ["('Giannis Antetokounmpo' in acquired or 'Jalen Brunson' in acquired)", "not ('Giannis Antetokounmpo' in acquired and 'Jalen Brunson' in acquired)", "('Karl-Anthony Towns' not in outgoing or 'Giannis Antetokounmpo' in acquired)"]
Treat "either X or Y, but not both" as TWO expressions: the OR and the NOT-AND.
For "trade X only if receiving Y", use "X not in outgoing or Y in acquired".
If the user requires both X and Y, use one AND expression (or equivalent
individual membership expressions).
Use acquired for players the user's team RECEIVES and outgoing for players
it SENDS. Avoid assumptions about other teams' player flows.
Do not invent constraints from preferred or merely allowable players.
Do not use player-type conditions, function calls, Python lambdas, or variables
besides acquired/outgoing. If no explicit mandatory condition exists, use [].

Distinguish REQUIRED from merely PREFERRED acquisitions. Explicitly naming
someone does not necessarily make acquisition mandatory: follow the wording.
Likewise, "willing to part with" means allowed, not necessarily preferred,
unless the user expresses a desire to move that player.
If a category is absent, put [] rather than speculating.

USER STATEMENT:
{user_statement}

All lists must use valid Python syntax with each player name, team name, or descriptive string enclosed in quotes. Example: Untouchable Players: ["Jalen Brunson", "OG Anunoby"]. Never output unquoted names.

Return only the labeled fields above. No explanations or Markdown.
"""
    )

    print("\nPARSED USER PREFERENCES:\n", demands.output_text)
    
    logic_match = re.search(r"(?im)^Required Trade Logic:\s*(\[[^\n]*\])\s*$", demands.output_text)
    if logic_match is None:
        raise ValueError("GPT omitted Required Trade Logic list")
    raw_logic = ast.literal_eval(logic_match.group(1))
    if not isinstance(raw_logic, list) or not all(isinstance(x, str) for x in raw_logic):
        raise ValueError("Required Trade Logic must be a list of Python expressions")

    
    
    def make_trade_rule(expression):
        tree = ast.parse(expression, mode="eval")

        def check_syntax(node):
            if isinstance(node, ast.Expression):
                return check_syntax(node.body)
            if isinstance(node, ast.BoolOp) and isinstance(node.op, (ast.And, ast.Or)):
                if len(node.values) < 2:
                    raise ValueError("Boolean group requires at least two members")
                for child in node.values:
                    check_syntax(child)
                return
            if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
                return check_syntax(node.operand)
            if (isinstance(node, ast.Compare) and len(node.ops) == 1
                    and isinstance(node.ops[0], (ast.In, ast.NotIn))
                    and isinstance(node.left, ast.Constant)
                    and isinstance(node.left.value, str)
                    and isinstance(node.comparators[0], ast.Name)
                    and node.comparators[0].id in {"acquired", "outgoing"}):
                return
            raise ValueError(f"Unsupported trade logic expression: {expression!r}")

        check_syntax(tree)

        def evaluate(node, acquired, outgoing):
            if isinstance(node, ast.Expression):
                return evaluate(node.body, acquired, outgoing)
            if isinstance(node, ast.BoolOp):
                results = (evaluate(n, acquired, outgoing) for n in node.values)
                return all(results) if isinstance(node.op, ast.And) else any(results)
            if isinstance(node, ast.UnaryOp):
                return not evaluate(node.operand, acquired, outgoing)
            selected = acquired if node.comparators[0].id == "acquired" else outgoing
            present = node.left.value in selected
            return present if isinstance(node.ops[0], ast.In) else not present

        def rule(acquired, outgoing):
            return evaluate(tree, set(acquired), set(outgoing))

        rule.expression = expression
        return rule

    required_logic = [make_trade_rule(expression) for expression in raw_logic]


    
    # STAGE 4 — GPT PLAYER-POOL REDUCTION AND PRIORITIZATION
    
    
    
    
    response = client.responses.create(
        model="gpt-5.6",
        input=f"""
You build four DIFFERENT player pools for an NBA trade search.

USER PREFERENCES:
{demands.output_text}

PLAYER PROFILES AND THEIR CURRENT TEAMS:
{all_results}

Using ONLY the supplied team rosters and player attributes/strengths/weaknesses,
produce EXACTLY FOUR Python dictionaries in the order specified below.
Each dictionary maps exact existing NBA team names to lists of exact existing
player names. No duplicates, fabricated names, or empty teams.

DICTIONARY 1: required_targets
- On OTHER teams, players the user MUST acquire, per explicit mandatory wording.
- Do NOT classify all named players as required automatically: distinguish an
  actual condition from a mere preference or interest.
- Honor excluded teams even if the user names a player on such a team, unless
  the user explicitly grants an exception for that team.
- A specifically required acquisition overrides general skill/weakness dislikes,
  but not an explicit team ban.

DICTIONARY 2: tradeable_targets
- On OTHER teams, players the user is interested in acquiring based on requested
  player types, skills, roles or explicit nonmandatory player interests.
- Only include players who reasonably match, using supplied scouting profiles.
- Respect excluded teams, explicitly unwanted players and unwanted player types.
- Do NOT duplicate players already in required_targets.
- Do not include everyone generally allowable: these must be actual targets.

DICTIONARY 3: wanted_to_trade_away
- From the USER'S OWN team ONLY, players the user expresses a genuine desire
  to trade away (not merely willingness to part with).
- Never include untouchables unless the user explicitly overrides that status.

DICTIONARY 4: can_trade_away
- From the USER'S OWN team ONLY, all other players the user is willing to trade.
- Exclude explicitly untouchable names and players matching protected types.
- Preserve explicit user permission to move a player even if they otherwise
  match a broad protected type (e.g. willing to part with Mikal Bridges but
  protecting other good shooters).
- Include named players the user is willing to part with, even if not eager.
- Start from the ENTIRE supplied user-team roster, not just players the user names.
- Remove only explicitly untouchable names, players matching explicitly protected
  types, and players excluded by an explicit ONLY-named-players restriction.
- If no protection or ONLY restriction was expressed, return the entire roster
  except players already in wanted_to_trade_away.
- Do NOT duplicate any player already in wanted_to_trade_away.

These four dictionaries must be pairwise DISJOINT. Put every explicitly required
OR-alternative into required_targets, even if no single member is mandatory;
the Required Trade Logic field controls acceptable combinations.
They describe USER
preferences only, NOT legal trade eligibility. Do not apply salary/CBA rules.
Do NOT use external knowledge or web searches. Do NOT change a player's team.

OUTPUT:
Return exactly four raw Python dict expressions (not a list or variable
assignments), in this strict order:
1. required_targets
2. tradeable_targets
3. wanted_to_trade_away
4. can_trade_away

Example output structure (illustrative names only):
{{"Opposing Team": ["Required Player"]}}
{{"Other Team": ["Interesting Player"]}}
{{"My Team": ["Player To Move"]}}
{{"My Team": ["Other Allowed Player"]}}

No headings, explanation, code fences, or text between dicts.
"""
    )

    def parse_four_dictionaries(text):
        try:
            tree = ast.parse(text.strip(), mode="exec")
        except SyntaxError as exc:
            raise ValueError(f"GPT returned invalid pool syntax:\n{text}") from exc

        values = []
        for node in tree.body:
            if not isinstance(node, ast.Expr):
                raise ValueError(f"Unexpected non-dictionary expression:\n{text}")
            value = ast.literal_eval(node.value)
            if not isinstance(value, dict):
                raise ValueError("Expected exactly four dictionaries")
            values.append(value)
        if len(values) != 4:
            raise ValueError(f"Expected four dictionaries, got {len(values)}:\n{text}")
        return values

    required_targets, tradeable_targets, wanted_to_trade_away, can_trade_away = (
        parse_four_dictionaries(response.output_text)
    )

    
    
    roster_lookup = {
        team: {normalize_name(name).casefold(): name for name in names}
        for team, names in league_players.items()
    }

    player_teams = {}
    for roster_team, roster_names in league_players.items():
        for roster_name in roster_names:
            player_teams.setdefault(normalize_name(roster_name).casefold(), []).append(
                (roster_team, roster_name)
            )

    user_team_match = re.search(r"(?im)^Team Name:\s*(.+?)\s*$", demands.output_text)
    named_user_team = user_team_match.group(1).strip()
    user_team_candidates = [
        name for name in league_players
        if name.casefold() == named_user_team.casefold()
        or name.casefold().endswith(" " + named_user_team.casefold())
    ]
    user_team_for_validation = user_team_candidates[0] if user_team_candidates else named_user_team

    def verify_roster_pool(pool, pool_name):
        clean = {}
        outgoing_pool = pool_name in {"wanted_to_trade_away", "can_trade_away"}
        for team, names in pool.items():
            if team not in roster_lookup:
                print(f"SKIPPED | {pool_name}: unknown team {team!r}")
                continue
            for name in names:
                key = normalize_name(name).casefold()
                actual = roster_lookup[team].get(key)
                assigned_team = team
                if actual is None:
                    matches = player_teams.get(key, [])
                    if len(matches) != 1:
                        print(f"SKIPPED | {pool_name}: {name!r} not verified on a unique scraped roster")
                        continue
                    assigned_team, actual = matches[0]
                    print(f"CORRECTED | {pool_name}: {name!r} moved from {team} to {assigned_team} using scraped rosters")
                if outgoing_pool and assigned_team != user_team_for_validation:
                    print(f"SKIPPED | {pool_name}: {actual!r} is not on the user's team")
                    continue
                if not outgoing_pool and assigned_team == user_team_for_validation:
                    print(f"SKIPPED | {pool_name}: {actual!r} is on the user's team")
                    continue
                clean.setdefault(assigned_team, set()).add(actual)
        return {team: sorted(names) for team, names in clean.items() if names}

    required_targets = verify_roster_pool(required_targets, "required_targets")
    tradeable_targets = verify_roster_pool(tradeable_targets, "tradeable_targets")
    wanted_to_trade_away = verify_roster_pool(wanted_to_trade_away, "wanted_to_trade_away")
    can_trade_away = verify_roster_pool(can_trade_away, "can_trade_away")

    user_team_match = re.search(r"(?im)^Team Name:\s*(.+?)\s*$", demands.output_text)
    user_team_raw = user_team_match.group(1).strip()
    if user_team_raw.startswith("[") and user_team_raw.endswith("]"):
        try:
            parsed_team = ast.literal_eval(user_team_raw)
            user_team_raw = str(parsed_team[0]).strip() if isinstance(parsed_team, list) and parsed_team else user_team_raw
        except (SyntaxError, ValueError):
            user_team_raw = user_team_raw[1:-1].strip().strip("\"'")
    user_team_raw = user_team_raw.strip().strip("\"'")
    possible_teams = [
        name for name in league_players
        if name.casefold() == user_team_raw.casefold()
        or name.casefold().endswith(" " + user_team_raw.casefold())
    ]
    user_team_name = possible_teams[0] if possible_teams else user_team_raw

    def preference_list(label):
        match = re.search(r"(?im)^" + re.escape(label) + r":\s*(\[[^\n]*\])\s*$", demands.output_text)
        if not match:
            return []
        raw = match.group(1).strip()
        try:
            values = ast.literal_eval(raw)
            return values if isinstance(values, list) else []
        except (ValueError, SyntaxError):
            inner = raw[1:-1].strip()
            if not inner:
                return []
            return [item.strip().strip("\"'") for item in inner.split(",") if item.strip()]

    untouchable_names = {normalize_name(name).casefold() for name in preference_list("Untouchable Players")}
    protected_types = preference_list("Untouchable Player Types")
    explicitly_allowed = {normalize_name(name).casefold() for name in preference_list("Other Allowed Outgoing Players")}
    explicitly_wanted = {normalize_name(name).casefold() for name in preference_list("Preferred Outgoing Players")}
    own_roster = league_players[user_team_name]
    wanted_names = {name for names in wanted_to_trade_away.values() for name in names}
    wanted_names.update(name for name in own_roster if normalize_name(name).casefold() in explicitly_wanted)
    wanted_names = {name for name in wanted_names if normalize_name(name).casefold() not in untouchable_names}
    wanted_to_trade_away = {user_team_name: sorted(wanted_names)} if wanted_names else {}

    only_named_outgoing = bool(re.search(r"\b(?:only|exclusively)\s+(?:trade|move|send|give up|part with)\s+(?:the\s+)?(?:following\s+)?(?:players|guys|people|names)\b|\b(?:trade|move|send)\s+only\b", user_statement, re.IGNORECASE))
    selected_by_gpt = set(can_trade_away.get(user_team_name, [])) | wanted_names
    if only_named_outgoing:
        baseline = selected_by_gpt | {name for name in own_roster if normalize_name(name).casefold() in explicitly_allowed}
    elif protected_types:
        baseline = selected_by_gpt | {name for name in own_roster if normalize_name(name).casefold() in explicitly_allowed}
    else:
        baseline = set(own_roster)
    allowed = sorted(name for name in baseline if name not in wanted_names and normalize_name(name).casefold() not in untouchable_names)
    can_trade_away = {user_team_name: allowed} if allowed else {}

    
    seen = set()
    def deduplicate_pool(pool):
        cleaned = {}
        for team, players in pool.items():
            unique = [name for name in players if (team, name) not in seen]
            seen.update((team, name) for name in unique)
            if unique:
                cleaned[team] = unique
        return cleaned

    required_targets = deduplicate_pool(required_targets)
    tradeable_targets = deduplicate_pool(tradeable_targets)
    wanted_to_trade_away = deduplicate_pool(wanted_to_trade_away)
    can_trade_away = deduplicate_pool(can_trade_away)

    print("\nREQUIRED TARGETS (BEFORE CBA):", required_targets)
    print("\nTRADEABLE TARGETS (BEFORE CBA):", tradeable_targets)
    print("\nWANTED TO TRADE AWAY (BEFORE CBA):", wanted_to_trade_away)
    print("\nCAN TRADE AWAY (BEFORE CBA):", can_trade_away)

    
    # STAGE 5 — INDIVIDUAL TRADE ELIGIBILITY (SPOTRAC ONLY)
    
    
    
    
    def filter_pool(pool, pool_name):
        remaining = {}
        removed = []
        for team, names in pool.items():
            accepted = []
            for player in names:
                restrictions = restriction_lookup.get((str(team), normalize_name(player).casefold()), [])
                individual = [restriction for restriction in restrictions if restriction["restriction_date"] > TRADE_DATE and not any(term in str(restriction["restriction_type"]).lower() for term in ("aggregate trade", "poison pill"))]
                if individual:
                    removed.append((team, player, "; ".join(f"{r["restriction_type"]} until {r["restriction_date"]}" for r in individual)))
                else:
                    accepted.append(player)
            if accepted:
                remaining[team] = sorted(set(accepted))
        print(f"\n{pool_name}: {sum(map(len, remaining.values()))} retained, {len(removed)} excluded")
        for team, player, reason in removed:
            print(f"  EXCLUDED | {team} | {player}: {reason}")
        return remaining

    eligible_required_targets = filter_pool(required_targets, "ELIGIBLE REQUIRED TARGETS")
    eligible_tradeable_targets = filter_pool(tradeable_targets, "ELIGIBLE TRADEABLE TARGETS")
    eligible_wanted_to_trade_away = filter_pool(wanted_to_trade_away, "ELIGIBLE WANTED TO TRADE AWAY")
    eligible_can_trade_away = filter_pool(can_trade_away, "ELIGIBLE CAN TRADE AWAY")

    print("\n" + "=" * 72)
    print("FINAL PLAYER DICTIONARIES — AFTER INDIVIDUAL ELIGIBILITY FILTERING")
    print("=" * 72)
    print("\nREQUIRED TARGETS:", eligible_required_targets)
    print("\nTRADEABLE TARGETS:", eligible_tradeable_targets)
    print("\nWANTED TO TRADE AWAY:", eligible_wanted_to_trade_away)
    print("\nCAN TRADE AWAY:", eligible_can_trade_away)
    print("\nREQUIRED TRADE LOGIC:", [rule.expression for rule in required_logic])

    # STAGE 6 — GENERATE EVERY TWO-TEAM PLAYER TRADE AND APPLY REQUIRED TRADE LOGIC


    outgoing_players = sorted(set(chain(
        eligible_wanted_to_trade_away.get(user_team_name, []),
        eligible_can_trade_away.get(user_team_name, []),
    )))

    def nonempty_subsets(players):
        for count in range(1, len(players) + 1):
            yield from combinations(players, count)

    number = 0
    legal_count = 0
    legal_trades = []
    for other_team in sorted(set(eligible_required_targets) | set(eligible_tradeable_targets)):
        if other_team == user_team_name:
            continue
        incoming_players = sorted(set(chain(
            eligible_required_targets.get(other_team, []),
            eligible_tradeable_targets.get(other_team, []),
        )))
        for sent_by_user in nonempty_subsets(outgoing_players):
            for sent_by_other in nonempty_subsets(incoming_players):
                number += 1
                legal = all(rule(set(sent_by_other), set(sent_by_user)) for rule in required_logic)
                print(f"TRADE {number}: {'LEGAL' if legal else 'ILLEGAL'}")
                if legal:
                    legal_count += 1
                    legal_trades.append((number, user_team_name, list(sent_by_user), other_team, list(sent_by_other)))
                    print(f"{user_team_name} sends: {list(sent_by_user)}")
                    print(f"{other_team} sends: {list(sent_by_other)}")

    print(f"Total trades: {number}; LEGAL (required logic passed): {legal_count}; ILLEGAL: {number - legal_count}")
    print("\nALL LEGAL TRADES")
    for trade_number, team_a, sent_a, team_b, sent_b in legal_trades:
        print(f"TRADE {trade_number}: LEGAL")
        print(f"{team_a} sends: {sent_a}")
        print(f"{team_b} sends: {sent_b}")
    return (eligible_required_targets, eligible_tradeable_targets,
            eligible_wanted_to_trade_away, eligible_can_trade_away, required_logic)

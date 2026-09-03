import re
import time
import requests
import urllib3
import pandas as pd
from io import StringIO
from datetime import date
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SPOTRAC_TEAMS = {
    "atlanta-hawks": ("Atlanta Hawks", "ATL"), "boston-celtics": ("Boston Celtics", "BOS"), "brooklyn-nets": ("Brooklyn Nets", "BKN"),
    "charlotte-hornets": ("Charlotte Hornets", "CHA"), "chicago-bulls": ("Chicago Bulls", "CHI"), "cleveland-cavaliers": ("Cleveland Cavaliers", "CLE"),
    "dallas-mavericks": ("Dallas Mavericks", "DAL"), "denver-nuggets": ("Denver Nuggets", "DEN"), "detroit-pistons": ("Detroit Pistons", "DET"),
    "golden-state-warriors": ("Golden State Warriors", "GSW"), "houston-rockets": ("Houston Rockets", "HOU"), "indiana-pacers": ("Indiana Pacers", "IND"),
    "la-clippers": ("Los Angeles Clippers", "LAC"), "los-angeles-lakers": ("Los Angeles Lakers", "LAL"), "memphis-grizzlies": ("Memphis Grizzlies", "MEM"),
    "miami-heat": ("Miami Heat", "MIA"), "milwaukee-bucks": ("Milwaukee Bucks", "MIL"), "minnesota-timberwolves": ("Minnesota Timberwolves", "MIN"),
    "new-orleans-pelicans": ("New Orleans Pelicans", "NOP"), "new-york-knicks": ("New York Knicks", "NYK"), "oklahoma-city-thunder": ("Oklahoma City Thunder", "OKC"),
    "orlando-magic": ("Orlando Magic", "ORL"), "philadelphia-76ers": ("Philadelphia 76ers", "PHI"), "phoenix-suns": ("Phoenix Suns", "PHX"),
    "portland-trail-blazers": ("Portland Trail Blazers", "POR"), "sacramento-kings": ("Sacramento Kings", "SAC"), "san-antonio-spurs": ("San Antonio Spurs", "SAS"),
    "toronto-raptors": ("Toronto Raptors", "TOR"), "utah-jazz": ("Utah Jazz", "UTA"), "washington-wizards": ("Washington Wizards", "WAS")
}

PICK_ALIASES = {
    "Atlanta Hawks": ["Atlanta"], "Boston Celtics": ["Boston"], "Brooklyn Nets": ["Brooklyn"], "Charlotte Hornets": ["Charlotte"],
    "Chicago Bulls": ["Chicago"], "Cleveland Cavaliers": ["Cleveland"], "Dallas Mavericks": ["Dallas"], "Denver Nuggets": ["Denver"],
    "Detroit Pistons": ["Detroit"], "Golden State Warriors": ["Golden State"], "Houston Rockets": ["Houston"], "Indiana Pacers": ["Indiana"],
    "Los Angeles Clippers": ["L.A. Clippers", "LA Clippers", "Los Angeles Clippers"], "Los Angeles Lakers": ["L.A. Lakers", "LA Lakers", "Los Angeles Lakers"],
    "Memphis Grizzlies": ["Memphis"], "Miami Heat": ["Miami"], "Milwaukee Bucks": ["Milwaukee"], "Minnesota Timberwolves": ["Minnesota"],
    "New Orleans Pelicans": ["New Orleans"], "New York Knicks": ["New York"], "Oklahoma City Thunder": ["Oklahoma City"], "Orlando Magic": ["Orlando"],
    "Philadelphia 76ers": ["Philadelphia"], "Phoenix Suns": ["Phoenix"], "Portland Trail Blazers": ["Portland"], "Sacramento Kings": ["Sacramento"],
    "San Antonio Spurs": ["San Antonio"], "Toronto Raptors": ["Toronto"], "Utah Jazz": ["Utah"], "Washington Wizards": ["Washington"]
}

MIN_CONTRACT_SEASON = 2026
TRANSACTION_START = "2021-07-01"
REQUEST_DELAY = 0.25
SAVE_EXCEL = True
PLAYER_POSITIONS = {"PG", "SG", "SF", "PF", "C", "G", "F", "G-F", "F-G", "F-C", "C-F"}
OUTPUT_FILE = "nba_trade_engine_data.xlsx"


def make_session():
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36", "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9", "Referer": "https://www.google.com/"})
    return session


def get_html(session, url, allow_insecure_fallback = False):
    try:
        response = session.get(url, timeout = 30)
    except requests.exceptions.SSLError:
        if not allow_insecure_fallback:
            raise
        print(f"    SSL verification failed for {url}. Retrying this public page with verify = False...")
        response = session.get(url, timeout = 30, verify = False)
    if response.status_code == 403:
        raise RuntimeError(f"403 Forbidden from {url}. The source may be blocking automated requests; rerun later or increase REQUEST_DELAY.")
    response.raise_for_status()
    return response.text


def flatten_columns(df):
    df = df.copy()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(x) for x in col if str(x) != "nan" and not str(x).startswith("Unnamed")]).strip() for col in df.columns]
    else:
        df.columns = [str(col).strip() for col in df.columns]
    return df


def read_tables(html):
    try:
        return [flatten_columns(df) for df in pd.read_html(StringIO(html))]
    except ValueError:
        return []


def money_to_number(value):
    if pd.isna(value):
        return pd.NA
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "—", "nan", "None"}:
        return pd.NA
    match = re.search(r"\$?(-?\d+(?:\.\d+)?)\s*([MBK])?", text, flags = re.I)
    if not match:
        return pd.NA
    number = float(match.group(1))
    suffix = (match.group(2) or "").upper()
    if suffix == "B":
        number *= 1_000_000_000
    elif suffix == "M":
        number *= 1_000_000
    elif suffix == "K":
        number *= 1_000
    return int(round(number))


def percent_to_number(value):
    if pd.isna(value):
        return pd.NA
    matches = re.findall(r"(-?\d+(?:\.\d+)?)%", str(value))
    return float(matches[-1]) if matches else pd.NA


def normalize_name(value):
    value = re.sub(r"\s+", " ", str(value)).strip()
    value = re.sub(r"\s+[†‡*]+$", "", value).strip()
    value = re.sub(r"\s*\([^)]{1,8}\)\s*$", "", value).strip()
    value = re.sub(r",\s*[A-Z/-]{1,6}$", "", value, flags = re.I).strip()
    tokens = value.split()
    for prefix_len in range(1, (len(tokens) // 2) + 1):
        if tokens[:prefix_len] == tokens[-prefix_len:] and len(tokens[prefix_len:]) >= 2:
            value = " ".join(tokens[prefix_len:])
            break
    return value


def season_start_year(label):
    match = re.match(r"^(20\d{2})-\d{2}$", str(label).strip())
    return int(match.group(1)) if match else None


def find_table(tables, required_columns):
    required_columns = [x.lower() for x in required_columns]
    for df in tables:
        columns = [str(col).lower() for col in df.columns]
        if all(any(required in col for col in columns) for required in required_columns):
            return df.copy()
    return None


def table_header_cells(table):
    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        texts = [re.sub(r"\s+", " ", cell.get_text(" ", strip = True)).strip() for cell in cells]
        if any(text.lower().startswith("player") for text in texts) and any(season_start_year(text) is not None for text in texts):
            return texts
    return []


def find_html_table(soup, required_headers):
    required_headers = [x.lower() for x in required_headers]
    for table in soup.find_all("table"):
        headers = [re.sub(r"\s+", " ", th.get_text(" ", strip = True)).strip() for th in table.find_all("th")]
        lower = [header.lower() for header in headers]
        if all(any(required in header for header in lower) for required in required_headers):
            return table
    return None


def find_table_after_heading(soup, heading_text, required_headers = None):
    heading = soup.find(lambda tag: tag.name in {"h1", "h2", "h3", "h4"} and heading_text.lower() in re.sub(r"\s+", " ", tag.get_text(" ", strip = True)).lower())
    if heading is not None:
        table = heading.find_next("table")
        if table is not None:
            if required_headers is None:
                return table
            headers = [re.sub(r"\s+", " ", th.get_text(" ", strip = True)).strip().lower() for th in table.find_all("th")]
            if all(any(required.lower() in header for header in headers) for required in required_headers):
                return table
    return find_html_table(soup, required_headers or [])


def clean_player_from_cell(cell):
    for anchor in cell.find_all("a"):
        href = str(anchor.get("href", ""))
        text = normalize_name(anchor.get_text(" ", strip = True))
        if text and ("/nba/player" in href or "/redirect/player" in href):
            return text
    return normalize_name(cell.get_text(" ", strip = True))


def get_spotrac_contract_terms(session):
    rows = []
    for i, (team_slug, (team_name, team_abbr)) in enumerate(SPOTRAC_TEAMS.items(), start = 1):
        print(f"[{i:02d}/30] Spotrac contracts: {team_name}")
        try:
            html = get_html(session, f"https://www.spotrac.com/nba/{team_slug}/contracts")
            soup = BeautifulSoup(html, "html.parser")
            table = find_html_table(soup, ["player", "start", "end", "yrs", "value", "aav"])
            if table is None:
                print("    Contract table not found.")
                continue
            headers = []
            header_row = None
            for tr in table.find_all("tr"):
                texts = [re.sub(r"\s+", " ", cell.get_text(" ", strip = True)).strip() for cell in tr.find_all(["th", "td"])]
                if any(text.lower().startswith("player") for text in texts) and any(text.lower() == "yrs" for text in texts):
                    headers = texts
                    header_row = tr
                    break
            if not headers:
                print("    Contract headers not found.")
                continue
            header_map = {re.sub(r"\s*\(\d+\)\s*$", "", header).strip(): idx for idx, header in enumerate(headers)}
            for tr in header_row.find_all_next("tr"):
                if tr.find_parent("table") != table:
                    break
                cells = tr.find_all("td")
                if len(cells) < max(6, len(headers) - 2):
                    continue
                name = clean_player_from_cell(cells[0])
                if not name or name.lower() in {"nan", "totals", "player"}:
                    continue
                def cell_text(header_name):
                    idx = header_map.get(header_name)
                    return cells[idx].get_text(" ", strip = True) if idx is not None and idx < len(cells) else pd.NA
                rows.append({"Name": name, "team": team_name, "team_abbreviation": team_abbr, "position": cell_text("Pos"), "signed_year": pd.to_numeric(cell_text("Start Year"), errors = "coerce"), "contract_type": cell_text("Type"), "age_at_signing": pd.to_numeric(cell_text("Age At Signing"), errors = "coerce"), "contract_start": pd.to_numeric(cell_text("Start"), errors = "coerce"), "contract_end": pd.to_numeric(cell_text("End"), errors = "coerce"), "contract_years": pd.to_numeric(cell_text("Yrs"), errors = "coerce"), "contract_value": money_to_number(cell_text("Value")), "contract_aav": money_to_number(cell_text("AAV")), "guaranteed_at_signing": money_to_number(cell_text("GTD @ Sign")), "practical_guaranteed": money_to_number(cell_text("Practical GTD"))})
        except Exception as e:
            print(f"    ERROR: {e}")
        time.sleep(REQUEST_DELAY)
    return pd.DataFrame(rows)


def parse_deadlines(tables):
    table = find_table(tables, ["Deadline Date", "Player", "Type", "Value"])
    rows = []
    if table is None:
        return pd.DataFrame(rows)
    for _, row in table.iterrows():
        name = normalize_name(row.get("Player"))
        type_text = str(row.get("Type", "")).strip()
        season_match = re.search(r"(20\d{2}-\d{2})", type_text)
        if not name or season_match is None:
            continue
        upper = type_text.upper()
        rows.append({"Name": name, "season": season_match.group(1), "player_option": upper.startswith("PLAYER ") and "PLAYER OPTION" in upper, "team_option": (upper.startswith("CLUB ") or upper.startswith("TEAM ")) and ("CLUB OPTION" in upper or "TEAM OPTION" in upper), "qualifying_offer": "RFA / QO" in upper or "QUALIFYING OFFER" in upper, "extension_eligible": "EXTENSION ELIGIBLE" in upper, "guarantee_deadline": "GUARANTEED" in upper, "deadline_date": pd.to_datetime(row.get("Deadline Date"), errors = "coerce"), "deadline_value": money_to_number(row.get("Value")), "deadline_type": type_text})
    return pd.DataFrame(rows)


def get_spotrac_yearly_contracts_and_cap(session):
    contract_rows = []
    cap_rows = []
    deadline_frames = []
    for i, (team_slug, (team_name, team_abbr)) in enumerate(SPOTRAC_TEAMS.items(), start = 1):
        print(f"[{i:02d}/30] Spotrac multi-year cap: {team_name}")
        try:
            html = get_html(session, f"https://www.spotrac.com/nba/{team_slug}/yearly/")
            soup = BeautifulSoup(html, "html.parser")
            tables = read_tables(html)
            deadlines = parse_deadlines(tables)
            if not deadlines.empty:
                deadlines["team"] = team_name
                deadlines["team_abbreviation"] = team_abbr
                deadline_frames.append(deadlines)

            active_table = find_table_after_heading(soup, "Active Roster", ["player", "pos", "age", "2026-27"])
            if active_table is not None:
                headers = table_header_cells(active_table)
                if headers:
                    header_row = None
                    for tr in active_table.find_all("tr"):
                        texts = [re.sub(r"\s+", " ", cell.get_text(" ", strip = True)).strip() for cell in tr.find_all(["th", "td"])]
                        if texts == headers:
                            header_row = tr
                            break
                    for tr in header_row.find_all_next("tr") if header_row is not None else []:
                        if tr.find_parent("table") != active_table:
                            break
                        cells = tr.find_all("td")
                        if len(cells) < len(headers):
                            continue
                        name = clean_player_from_cell(cells[0])
                        if not name or name.lower() in {"nan", "totals", "player"}:
                            continue
                        position = cells[1].get_text(" ", strip = True) if len(cells) > 1 else pd.NA
                        age = pd.to_numeric(cells[2].get_text(" ", strip = True), errors = "coerce") if len(cells) > 2 else pd.NA
                        for idx, column in enumerate(headers):
                            year = season_start_year(column)
                            if year is None or year < MIN_CONTRACT_SEASON or idx >= len(cells):
                                continue
                            raw = re.sub(r"\s+", " ", cells[idx].get_text(" ", strip = True)).strip()
                            if raw in {"", "nan", "-", "—"} or "UFA" in raw.upper() or "RFA" in raw.upper():
                                continue
                            salary = money_to_number(raw)
                            if pd.isna(salary):
                                continue
                            contract_rows.append({"Name": name, "team": team_name, "team_abbreviation": team_abbr, "position": position, "age": age, "season": column, "season_start": year, "salary": salary, "cap_pct": percent_to_number(raw), "player_option": False, "team_option": False, "qualifying_offer": False, "extension_eligible": "ext. elig" in raw.lower(), "salary_status_raw": raw})

            summary = None
            for table in tables:
                first_col = table.iloc[:, 0].astype(str).str.strip() if not table.empty else pd.Series(dtype = str)
                if first_col.str.contains("Cap Maximum", case = False, na = False).any() and first_col.str.contains("1st Apron", case = False, na = False).any():
                    summary = table.copy()
                    break
            if summary is not None:
                label_col = summary.columns[0]
                labels = summary[label_col].astype(str).str.strip().tolist()
                def row_values(label, start_index = 0):
                    for idx in range(start_index, len(labels)):
                        if labels[idx].lower() == label.lower():
                            return idx, summary.iloc[idx]
                    return None, None
                _, cap_max_row = row_values("Cap Maximum")
                _, active_row = row_values("Active Cap")
                _, dead_row = row_values("Dead Cap")
                _, holds_row = row_values("Cap Holds")
                _, total_row = row_values("Total Cap Allocations")
                _, space_row = row_values("Cap Space")
                first_header = next((idx for idx, label in enumerate(labels) if "1st Apron" in label and "Space" not in label), None)
                second_header = next((idx for idx, label in enumerate(labels) if "2nd Apron" in label and "Space" not in label), None)
                _, first_threshold_row = row_values("Threshold", first_header + 1 if first_header is not None else 0)
                _, first_alloc_row = row_values("Allocations", first_header + 1 if first_header is not None else 0)
                _, first_space_row = row_values("1st Apron Space", first_header + 1 if first_header is not None else 0)
                _, second_threshold_row = row_values("Threshold", second_header + 1 if second_header is not None else 0)
                _, second_alloc_row = row_values("Allocations", second_header + 1 if second_header is not None else 0)
                _, second_space_row = row_values("2nd Apron Space", second_header + 1 if second_header is not None else 0)
                for column in summary.columns[1:]:
                    year = season_start_year(column)
                    if year is None or year < MIN_CONTRACT_SEASON:
                        continue
                    cap_space = money_to_number(space_row.get(column)) if space_row is not None else pd.NA
                    first_space = money_to_number(first_space_row.get(column)) if first_space_row is not None else pd.NA
                    second_space = money_to_number(second_space_row.get(column)) if second_space_row is not None else pd.NA
                    cap_rows.append({"team": team_name, "team_abbreviation": team_abbr, "season": column, "season_start": year, "salary_cap": money_to_number(cap_max_row.get(column)) if cap_max_row is not None else pd.NA, "active_cap": money_to_number(active_row.get(column)) if active_row is not None else pd.NA, "dead_cap": money_to_number(dead_row.get(column)) if dead_row is not None else pd.NA, "cap_holds": money_to_number(holds_row.get(column)) if holds_row is not None else pd.NA, "total_cap_allocations": money_to_number(total_row.get(column)) if total_row is not None else pd.NA, "cap_space": cap_space, "first_apron": money_to_number(first_threshold_row.get(column)) if first_threshold_row is not None else pd.NA, "first_apron_allocations": money_to_number(first_alloc_row.get(column)) if first_alloc_row is not None else pd.NA, "first_apron_space": first_space, "second_apron": money_to_number(second_threshold_row.get(column)) if second_threshold_row is not None else pd.NA, "second_apron_allocations": money_to_number(second_alloc_row.get(column)) if second_alloc_row is not None else pd.NA, "second_apron_space": second_space, "above_cap": bool(cap_space < 0) if not pd.isna(cap_space) else pd.NA, "above_first_apron": bool(first_space < 0) if not pd.isna(first_space) else pd.NA, "above_second_apron": bool(second_space < 0) if not pd.isna(second_space) else pd.NA})
        except Exception as e:
            print(f"    ERROR: {e}")
        time.sleep(REQUEST_DELAY)

    contracts_df = pd.DataFrame(contract_rows)
    deadlines_df = pd.concat(deadline_frames, ignore_index = True) if deadline_frames else pd.DataFrame()
    if not contracts_df.empty and not deadlines_df.empty:
        flag_cols = ["player_option", "team_option", "qualifying_offer", "extension_eligible", "guarantee_deadline"]
        grouped = deadlines_df.groupby(["Name", "team", "team_abbreviation", "season"], as_index = False)[flag_cols].max()
        contracts_df = contracts_df.merge(grouped, on = ["Name", "team", "team_abbreviation", "season"], how = "left", suffixes = ("", "_deadline"))
        for col in ["player_option", "team_option", "qualifying_offer", "extension_eligible"]:
            deadline_col = f"{col}_deadline"
            if deadline_col in contracts_df.columns:
                contracts_df[col] = contracts_df[col].fillna(False) | contracts_df[deadline_col].fillna(False)
                contracts_df = contracts_df.drop(columns = deadline_col)
        if "guarantee_deadline" in contracts_df.columns:
            contracts_df["guarantee_deadline"] = contracts_df["guarantee_deadline"].fillna(False).astype(bool)
    return contracts_df, pd.DataFrame(cap_rows), deadlines_df


def get_spotrac_trade_exceptions(session):
    rows = []
    for i, (team_slug, (team_name, team_abbr)) in enumerate(SPOTRAC_TEAMS.items(), start = 1):
        print(f"[{i:02d}/30] Spotrac trade exceptions: {team_name}")
        try:
            html = get_html(session, f"https://www.spotrac.com/nba/{team_slug}/cap/_/year/{MIN_CONTRACT_SEASON}")
            table = find_table(read_tables(html), ["Type", "Reason", "Expires", "Original", "Available"])
            if table is None:
                continue
            for _, row in table.iterrows():
                exception_type = str(row.get("Type", "")).strip()
                if exception_type.lower() != "trade":
                    continue
                rows.append({"team": team_name, "team_abbreviation": team_abbr, "season": f"{MIN_CONTRACT_SEASON}-{str(MIN_CONTRACT_SEASON + 1)[-2:]}", "exception_type": exception_type, "reason": row.get("Reason"), "used_on": row.get("Used On"), "expires": pd.to_datetime(row.get("Expires"), errors = "coerce"), "original_amount": money_to_number(row.get("Original")), "available_amount": money_to_number(row.get("Available"))})
        except Exception as e:
            print(f"    ERROR: {e}")
        time.sleep(REQUEST_DELAY)
    return pd.DataFrame(rows)


def classify_transaction(description):
    text = str(description).lower()
    if "extension" in text or "extended" in text:
        return "Extension"
    if "traded" in text or "trade to" in text or "trade from" in text:
        return "Trade"
    if "signed" in text or "re-signed" in text:
        return "Signing"
    return None


def parse_spotrac_transaction_page(html, source_team, source_abbr):
    soup = BeautifulSoup(html, "html.parser")
    date_pattern = re.compile(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+20\d{2}\b")
    keyword_pattern = re.compile(r"\b(?:signed|re-signed|extension|extended|traded)\b", flags = re.I)
    containers = []
    for tag in soup.find_all("li"):
        text = " ".join(tag.stripped_strings)
        if date_pattern.search(text) and keyword_pattern.search(text):
            containers.append(tag)
    if not containers:
        candidates = []
        for tag in soup.find_all("div"):
            text = " ".join(tag.stripped_strings)
            if date_pattern.search(text) and keyword_pattern.search(text) and 20 <= len(text) <= 1200:
                child_match = any(date_pattern.search(" ".join(child.stripped_strings)) and keyword_pattern.search(" ".join(child.stripped_strings)) for child in tag.find_all("div", recursive = False))
                if not child_match:
                    candidates.append(tag)
        containers = candidates
    rows = []
    seen = set()
    for container in containers:
        text = re.sub(r"\s+", " ", " ".join(container.stripped_strings)).strip()
        date_match = date_pattern.search(text)
        if not date_match:
            continue
        player_links = [a for a in container.find_all("a") if "/nba/player" in str(a.get("href", "")) or "/redirect/player" in str(a.get("href", ""))]
        if player_links:
            name = normalize_name(player_links[0].get_text(" ", strip = True))
        else:
            after_date = text[date_match.end():].strip()
            name_match = re.match(r"([^,]+?)(?:,\s*[A-Z]{1,3}|\s*\([A-Z/-]{1,6}\))", after_date)
            name = normalize_name(name_match.group(1)) if name_match else normalize_name(text[:date_match.start()])
        if not name or name.lower() == "nan":
            continue
        position_match = re.search(re.escape(name) + r"\s*(?:,\s*|\()([A-Z/-]{1,6})\)?", text, flags = re.I)
        position = position_match.group(1).upper() if position_match else pd.NA
        if not pd.isna(position) and position not in PLAYER_POSITIONS:
            continue
        description = text[date_match.end():].strip().lstrip("-").strip()
        description = re.sub(r"^" + re.escape(name) + r"\s*(?:,\s*[A-Z/-]{1,6}|\s*\([A-Z/-]{1,6}\))?\s*", "", description, flags = re.I)
        description = re.sub(r"^(Signing|Trade|Extension)\s+\1\s+", r"\1 ", description, flags = re.I)
        if "pending" in description.lower():
            continue
        transaction_type = classify_transaction(description)
        if transaction_type is None:
            continue
        transaction_date = pd.to_datetime(date_match.group(0), errors = "coerce")
        key = (name, transaction_date, transaction_type, description)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"Name": name, "position": position, "date": transaction_date, "transaction_type": transaction_type, "team": source_team, "team_abbreviation": source_abbr, "description": description})
    return rows


def get_spotrac_transactions(session):
    rows = []
    end_date = date.today().isoformat()
    for i, (_, (team_name, team_abbr)) in enumerate(SPOTRAC_TEAMS.items(), start = 1):
        print(f"[{i:02d}/30] Spotrac transactions: {team_name}")
        try:
            url = f"https://www.spotrac.com/nba/transactions/_/start/{TRANSACTION_START}/end/{end_date}/team/{team_abbr.lower()}"
            rows.extend(parse_spotrac_transaction_page(get_html(session, url), team_name, team_abbr))
        except Exception as e:
            print(f"    ERROR: {e}")
        time.sleep(REQUEST_DELAY)
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.drop_duplicates(subset = ["Name", "date", "transaction_type", "description"]).sort_values(["date", "Name"], ascending = [False, True]).reset_index(drop = True)
    return df


def pick_is_referenced(traded_df, team_name, draft_year, draft_round):
    round_terms = "(?:1st|first)" if draft_round == 1 else "(?:2nd|second)"
    combined = (traded_df["asset"].fillna("") + " " + traded_df["details"].fillna("")).astype(str)
    for alias in PICK_ALIASES.get(team_name, [team_name]):
        possessive = re.escape(alias) + r"(?:'s|')"
        pattern = re.compile(rf"{possessive}\s+{draft_year}\s+{round_terms}\s+round\s+pick", flags = re.I)
        if combined.str.contains(pattern, regex = True, na = False).any():
            return True
    return False


def get_courtside_draft_assets(session):
    print("Pulling Courtside News future draft assets...")
    url = "https://courtsidenews.com/nba/draft/picks"
    html = get_html(session, url, allow_insecure_fallback = True)
    soup = BeautifulSoup(html, "html.parser")
    strings = [re.sub(r"\s+", " ", text).strip() for text in soup.stripped_strings]
    team_lookup = {team_name: (team_name, team_abbr) for _, (team_name, team_abbr) in SPOTRAC_TEAMS.items()}
    team_lookup["LA Clippers"] = ("Los Angeles Clippers", "LAC")
    team_names = set(team_lookup)
    rows = []
    current_team = None
    current_team_abbr = None
    current_year = None
    current_direction = None
    current_row = None
    started = False

    def finish_current():
        nonlocal current_row
        if current_row is None:
            return
        details = " ".join(current_row.pop("detail_parts", [])).strip()
        current_row["details"] = details if details else pd.NA
        combined = f"{current_row.get('asset', '')} {details}".lower()
        current_row["is_swap"] = "swap" in combined or "right to swap" in combined
        current_row["is_protected"] = "protect" in combined
        current_row["is_conditional"] = any(term in combined for term in ["if ", "more favorable", "less favorable", "most favorable", "least favorable", "convey", "extinguished", "condition", "outgoing"])
        rows.append(current_row)
        current_row = None

    for text in strings:
        if text == "NBA Future Draft Picks Tracker":
            started = True
            continue
        if not started:
            continue
        if text == "NBA Standings":
            finish_current()
            break
        if text in team_names:
            finish_current()
            current_team, current_team_abbr = team_lookup[text]
            current_year = None
            current_direction = None
            continue
        if current_team is None:
            continue
        if re.fullmatch(r"20\d{2}", text):
            year = int(text)
            if year >= MIN_CONTRACT_SEASON + 1:
                finish_current()
                current_year = year
                current_direction = None
            continue
        if text in {"IN", "OUT"}:
            finish_current()
            current_direction = text
            continue
        asset_match = re.match(r"^(20\d{2}) (first|second) round draft pick\b", text, flags = re.I)
        if asset_match and current_year is not None and current_direction is not None:
            finish_current()
            asset = re.sub(r"\s*Protections and trade notes\s*$", "", text, flags = re.I).strip()
            current_row = {"team": current_team, "team_abbreviation": current_team_abbr, "draft_year": int(asset_match.group(1)), "round": 1 if asset_match.group(2).lower() == "first" else 2, "direction": current_direction, "asset": asset, "detail_parts": []}
            continue
        if current_row is not None and text.lower() != "protections and trade notes":
            current_row["detail_parts"].append(text)

    finish_current()
    traded_df = pd.DataFrame(rows)
    if traded_df.empty:
        raise RuntimeError("No future draft-pick data was parsed from Courtside News.")
    traded_df = traded_df[traded_df["draft_year"] >= MIN_CONTRACT_SEASON + 1].copy()

    own_rows = []
    max_draft_year = max(int(traded_df["draft_year"].max()), MIN_CONTRACT_SEASON + 7)
    for _, (team_name, team_abbr) in SPOTRAC_TEAMS.items():
        for draft_year in range(MIN_CONTRACT_SEASON + 1, max_draft_year + 1):
            for draft_round in [1, 2]:
                if not pick_is_referenced(traded_df, team_name, draft_year, draft_round):
                    round_name = "first" if draft_round == 1 else "second"
                    own_rows.append({"team": team_name, "team_abbreviation": team_abbr, "draft_year": draft_year, "round": draft_round, "direction": "OWN", "asset": f"{draft_year} {round_name} round own draft pick", "details": "Unencumbered own pick; this original pick is not referenced in any traded-pick obligation, protection, or swap in the tracker.", "is_swap": False, "is_protected": False, "is_conditional": False})

    own_df = pd.DataFrame(own_rows)
    draft_assets_df = pd.concat([traded_df, own_df], ignore_index = True, sort = False)
    direction_order = pd.CategoricalDtype(["OWN", "IN", "OUT"], ordered = True)
    draft_assets_df["direction"] = draft_assets_df["direction"].astype(direction_order)
    draft_assets_df = draft_assets_df.sort_values(["team", "draft_year", "round", "direction", "asset"]).reset_index(drop = True)
    draft_assets_df["direction"] = draft_assets_df["direction"].astype(str)
    return draft_assets_df


def attach_contract_terms(annual_contracts_df, contract_terms_df):
    if annual_contracts_df.empty or contract_terms_df.empty:
        return annual_contracts_df.copy()
    annual = annual_contracts_df.copy().reset_index(drop = True)
    terms = contract_terms_df.copy()
    term_columns = ["signed_year", "contract_type", "age_at_signing", "contract_start", "contract_end", "contract_years", "contract_value", "contract_aav", "guaranteed_at_signing", "practical_guaranteed"]
    output_rows = []
    for _, row in annual.iterrows():
        candidates = terms[(terms["Name"] == row["Name"]) & (terms["team_abbreviation"] == row["team_abbreviation"])].copy()
        season_year = int(row["season_start"])
        if not candidates.empty:
            valid = candidates[(pd.to_numeric(candidates["contract_start"], errors = "coerce") <= season_year) & (pd.to_numeric(candidates["contract_end"], errors = "coerce") >= season_year)]
            if valid.empty:
                valid = candidates[pd.to_numeric(candidates["contract_start"], errors = "coerce") <= season_year]
            if valid.empty:
                valid = candidates
            valid = valid.assign(_start = pd.to_numeric(valid["contract_start"], errors = "coerce")).sort_values("_start", ascending = False)
            chosen = valid.iloc[0]
            for col in term_columns:
                row[col] = chosen.get(col, pd.NA)
        else:
            for col in term_columns:
                row[col] = pd.NA
        output_rows.append(row.to_dict())
    return pd.DataFrame(output_rows)


def add_transaction_dates_to_contracts(contracts_df, transactions_df):
    if contracts_df.empty or transactions_df.empty:
        return contracts_df
    transactions = transactions_df.copy()
    transactions["Name"] = transactions["Name"].map(normalize_name)
    transactions["event_year"] = pd.to_datetime(transactions["date"], errors = "coerce").dt.year
    contracts = contracts_df.copy()
    signed_dates = []
    signed_types = []
    signed_descriptions = []
    latest_trade_dates = []
    latest_trade_descriptions = []
    for _, contract in contracts.iterrows():
        name = normalize_name(contract["Name"])
        signed_year = pd.to_numeric(contract.get("signed_year"), errors = "coerce")
        events = transactions[(transactions["Name"] == name) & (transactions["transaction_type"].isin(["Signing", "Extension"]))].copy()
        if not pd.isna(signed_year):
            same_year = events[events["event_year"] == int(signed_year)]
            if not same_year.empty:
                events = same_year
        if not events.empty:
            event = events.sort_values("date").iloc[-1]
            signed_dates.append(event["date"])
            signed_types.append(event["transaction_type"])
            signed_descriptions.append(event["description"])
        else:
            signed_dates.append(pd.NaT)
            signed_types.append(pd.NA)
            signed_descriptions.append(pd.NA)
        trades = transactions[(transactions["Name"] == name) & (transactions["transaction_type"] == "Trade")].copy()
        if not trades.empty:
            trade = trades.sort_values("date").iloc[-1]
            latest_trade_dates.append(trade["date"])
            latest_trade_descriptions.append(trade["description"])
        else:
            latest_trade_dates.append(pd.NaT)
            latest_trade_descriptions.append(pd.NA)
    contracts["contract_signed_date"] = signed_dates
    contracts["contract_event_type"] = signed_types
    contracts["contract_event_description"] = signed_descriptions
    contracts["latest_trade_date"] = latest_trade_dates
    contracts["latest_trade_description"] = latest_trade_descriptions
    return contracts


def validate_data(contracts_df, team_cap_df, draft_assets_df):
    errors = []
    if contracts_df.empty:
        errors.append("contracts_df is empty")
    else:
        bad_salary = contracts_df[pd.to_numeric(contracts_df["salary"], errors = "coerce") > 120_000_000]
        if not bad_salary.empty:
            errors.append(f"{len(bad_salary)} contract rows have salary > $120M; salary/percentage parsing likely failed")
        bad_pct = contracts_df[pd.to_numeric(contracts_df["cap_pct"], errors = "coerce") > 60]
        if not bad_pct.empty:
            errors.append(f"{len(bad_pct)} contract rows have cap_pct > 60%; percentage parsing likely failed")
        current = contracts_df[contracts_df["season_start"] == MIN_CONTRACT_SEASON]
        counts = current.groupby("team_abbreviation")["Name"].nunique()
        missing_or_tiny = [abbr for _, (_, abbr) in SPOTRAC_TEAMS.items() if counts.get(abbr, 0) < 5]
        if missing_or_tiny:
            errors.append(f"Current-season contract pull returned fewer than 5 players for: {', '.join(missing_or_tiny)}")
    cap_current = team_cap_df[team_cap_df["season_start"] == MIN_CONTRACT_SEASON] if not team_cap_df.empty else pd.DataFrame()
    if cap_current.empty or cap_current["team_abbreviation"].nunique() != 30:
        errors.append("Current-season cap table does not contain all 30 teams")
    if draft_assets_df.empty:
        errors.append("draft_assets_df is empty")
    if errors:
        raise RuntimeError("DATA VALIDATION FAILED:\n- " + "\n- ".join(errors))
    print("Data validation passed.")


def get_nba_trade_engine_data():
    session = make_session()
    draft_assets_df = get_courtside_draft_assets(session)
    print(f"Draft assets successfully pulled: {len(draft_assets_df):,} rows\n")
    contract_terms_df = get_spotrac_contract_terms(session)
    annual_contracts_df, team_cap_df, deadlines_df = get_spotrac_yearly_contracts_and_cap(session)
    trade_exceptions_df = get_spotrac_trade_exceptions(session)
    transactions_df = get_spotrac_transactions(session)
    contracts_df = attach_contract_terms(annual_contracts_df, contract_terms_df) if not annual_contracts_df.empty else contract_terms_df.copy()
    contracts_df = add_transaction_dates_to_contracts(contracts_df, transactions_df)
    if not contracts_df.empty and "season" in contracts_df.columns:
        contracts_df = contracts_df.drop_duplicates(subset = ["Name", "team", "season"]).sort_values(["team", "Name", "season"]).reset_index(drop = True)
    validate_data(contracts_df, team_cap_df, draft_assets_df)
    return contracts_df, transactions_df, team_cap_df, trade_exceptions_df, draft_assets_df, deadlines_df


def save_team_workbook(contracts_df, transactions_df, team_cap_df, trade_exceptions_df, draft_assets_df, deadlines_df, output_file = OUTPUT_FILE):
    sections = [("CONTRACTS", contracts_df), ("TRANSACTIONS", transactions_df), ("TEAM CAP / APRONS", team_cap_df), ("TRADE EXCEPTIONS", trade_exceptions_df), ("CONTRACT DEADLINES", deadlines_df), ("DRAFT ASSETS", draft_assets_df)]
    with pd.ExcelWriter(output_file, engine = "openpyxl", datetime_format = "yyyy-mm-dd") as writer:
        for _, (team_name, team_abbr) in SPOTRAC_TEAMS.items():
            sheet_name = team_name[:31]
            start_row = 0
            section_title_rows = []
            for section_name, df in sections:
                section_title_rows.append(start_row + 1)
                pd.DataFrame([[section_name]]).to_excel(writer, sheet_name = sheet_name, startrow = start_row, index = False, header = False)
                start_row += 1
                if df.empty:
                    team_df = pd.DataFrame()
                elif "team_abbreviation" in df.columns:
                    team_df = df[df["team_abbreviation"] == team_abbr].copy()
                elif "team" in df.columns:
                    team_df = df[df["team"] == team_name].copy()
                else:
                    team_df = pd.DataFrame()
                if team_df.empty:
                    pd.DataFrame({"Info": ["No data found"]}).to_excel(writer, sheet_name = sheet_name, startrow = start_row, index = False)
                    start_row += 3
                else:
                    team_df.to_excel(writer, sheet_name = sheet_name, startrow = start_row, index = False)
                    start_row += len(team_df) + 3
            worksheet = writer.book[sheet_name]
            worksheet.freeze_panes = "A2"
            for row_number in section_title_rows:
                cell = worksheet.cell(row = row_number, column = 1)
                cell.font = cell.font.copy(bold = True, size = 12)
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, pd.Timestamp):
                        cell.number_format = "yyyy-mm-dd"
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 40)
    print(f"Saved Excel workbook: {output_file}")


contracts_df, transactions_df, team_cap_df, trade_exceptions_df, draft_assets_df, deadlines_df = get_nba_trade_engine_data()

print("\nDONE")
print(f"Contracts: {len(contracts_df):,} rows")
print(f"Transactions: {len(transactions_df):,} rows")
print(f"Team cap/apron: {len(team_cap_df):,} rows")
print(f"Trade exceptions: {len(trade_exceptions_df):,} rows")
print(f"Contract deadlines: {len(deadlines_df):,} rows")
print(f"Draft assets: {len(draft_assets_df):,} rows")

if SAVE_EXCEL:
    save_team_workbook(contracts_df, transactions_df, team_cap_df, trade_exceptions_df, draft_assets_df, deadlines_df)

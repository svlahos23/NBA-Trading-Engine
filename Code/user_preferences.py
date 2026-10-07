from typing import Any, cast

import json
import httpx
from openai import OpenAI


user_statement = """I work for the New York Knicks. I am looking for a strong backup center to replace Mitchell Robinson. I also would like to trade for Payton Pritchard.
I am willing to part with Mikal Bridges. Apart from him, I am not willing to part with any other players who are decent shooters. I am not willing to part with Karl-Anthony Towns or Jalen Brunson.
I want to avoid trading for players that are bad shooters or are chokers. No James Harden. I do not want to trade with the Celtics.
"""


league_players = {
    'Atlanta Hawks': [
        'Aaron Wiggins',
        'Asa Newell',
        'Buddy Hield',
        'C.J. McCollum',
        'Corey Kispert',
        'Devin Carter',
        'Dyson Daniels',
        'Henri Veesaar',
        'Jalen Johnson',
        'Jock Landale',
        'Kingston Flemings',
        'Luguentz Dort',
        'Mouhamed Gueye',
        'Nickeil Alexander-Walker',
        'Onyeka Okongwu',
        'Ryan Nembhard',
        'Zuby Ejiofor'
    ],

    'Boston Celtics': [
        'Baylor Scheierman',
        'Chris Cenac Jr.',
        'Derrick White',
        'Hugo Gonzalez',
        'Jayson Tatum',
        'Jordan Walsh',
        'Luka Garza',
        'Mike Conley',
        'Milos Uzan',
        'Mitchell Robinson',
        'Neemias Queta',
        'Paul George',
        'Payton Pritchard',
        'Ron Harper Jr.',
        'Sam Hauser',
        'Tucker DeVries'
    ],

    'Brooklyn Nets': [
        'Ben Saraf',
        'Danny Wolf',
        "Day'Ron Sharpe",
        'Drake Powell',
        'Egor Demin',
        'Josh Minott',
        'Joshua Jefferson',
        'Julius Randle',
        'Keon Ellis',
        'Michael Porter Jr.',
        'Mikel Brown Jr.',
        'Moritz Wagner',
        'Noah Clowney',
        'Nolan Traore',
        'Terance Mann'
    ],

    'Charlotte Hornets': [
        'Brandon Miller',
        'Christian Anderson',
        'Coby White',
        'Dennis Schröder',
        'Dorian Finney-Smith',
        'Grant Williams',
        'Grayson Allen',
        'Hannes Steinbach',
        'Jarkel Joiner',
        'Kon Knueppel',
        'Liam McNeeley',
        'Moussa Diabate',
        'Naz Reid',
        'Pat Connaughton',
        "Royce O'Neale",
        'Ryan Kalkbrenner',
        'Sion James',
        'Tidjane Salaun'
    ],

    'Chicago Bulls': [
        'Caleb Wilson',
        'Dailyn Swain',
        'Isaac Okoro',
        'Jalen Smith',
        'Josh Giddey',
        'Leonard Miller',
        'Matas Buzelis',
        'Nicolas Claxton',
        'Noa Essengue',
        'Norman Powell',
        'Patrick Williams',
        'Rob Dillingham',
        'Tre Jones',
        'Yuki Kawamura',
        'Zach Collins'
    ],

    'Cleveland Cavaliers': [
        'Craig Porter Jr.',
        'Donovan Mitchell',
        'Evan Mobley',
        'Jarrett Allen',
        'Jaylon Tyson',
        'Khalifa Diop',
        'Mario Hezonja',
        'Meleek Thomas',
        "Nae'qwan Tomlin",
        'Peyton Watson',
        'Sam Merrill',
        'Thomas Bryant',
        'Tyrese Proctor'
    ],

    'Dallas Mavericks': [
        'Caleb Martin',
        'Cooper Flagg',
        'Daniel Gafford',
        'Dereck Lively II',
        'Kyrie Irving',
        'Marcus Sasser',
        'Max Christie',
        'Morez Johnson Jr.',
        'Moussa Cisse',
        'Naji Marshall',
        'PJ Washington',
        'Santi Aldama',
        'Sergio de Larrea',
        'Tarik Biberovic',
        'Zaccharie Risacher'
    ],

    'Denver Nuggets': [
        'Aaron Gordon',
        'Alpha Diallo',
        'Cameron Johnson',
        'Christian Braun',
        'DaRon Holmes II',
        'Jamal Murray',
        'Julian Strawther',
        'Marvin Bagley III',
        'Nikola Jokic',
        'Spencer Jones',
        'Trevon Brazile',
        'Tyus Jones',
        'Zeke Nnaji'
    ],

    'Detroit Pistons': [
        'Ausar Thompson',
        'Cade Cunningham',
        'Chaz Lanier',
        'Daniss Jenkins',
        'Duncan Robinson',
        'Ebuka Okorie',
        'Gary Harris',
        'Isaiah Joe',
        'Javonte Green',
        'John Collins',
        'Kevin Huerter',
        'Paul Reed',
        'Ron Holland II',
        'Taurean Prince',
        'Tolu Smith III'
    ],

    'Golden State Warriors': [
        'Al Horford',
        'Brandin Podziemski',
        'Brandon Williams',
        'Charles Bassey',
        'Dalen Terry',
        "De'Anthony Melton",
        'Draymond Green',
        'Gary Payton II',
        'Georges Niang',
        'Gui Santos',
        'Jimmy Butler',
        'Kristaps Porzingis',
        'Moses Moody',
        'Stephen Curry',
        'Will Richard',
        'Yaxel Lendeborg'
    ],

    'Houston Rockets': [
        'Alperen Sengun',
        'Amen Thompson',
        'Bogdan Bogdanovic',
        'Bruce Thornton',
        'Clint Capela',
        'Fred VanVleet',
        'Isaiah Crawford',
        'Jabari Smith Jr.',
        'Jae’Sean Tate',
        'Julian Phillips',
        'Kevin Durant',
        'Marcus Smart',
        'Oscar Tshiebwe',
        'Reed Sheppard',
        'Steven Adams',
        'Tari Eason'
    ],

    'Indiana Pacers': [
        'Aaron Nesmith',
        'Andrew Nembhard',
        'Ben Sheppard',
        'Ivica Zubac',
        'Jarace Walker',
        'Jay Huff',
        'Johnny Furphy',
        'Kelly Oubre Jr.',
        'Larry Nance Jr.',
        'Obi Toppin',
        'Pascal Siakam',
        'Quenton Jackson',
        'T.J. McConnell',
        'Tyrese Haliburton'
    ],

    'Los Angeles Clippers': [
        'Baba Miller',
        'Bradley Beal',
        'Brook Lopez',
        'Cameron Christie',
        'Darius Garland',
        'Derrick Jones Jr.',
        'Isaiah Jackson',
        'Johni Broome',
        'Jordan Miller',
        'Kawhi Leonard',
        'Keaton Wagler',
        'Kobe Sanders',
        'Kris Dunn',
        'Max Strus',
        'Rui Hachimura',
        'Yanic Konan Niederhauser'
    ],

    'Los Angeles Lakers': [
        'Adou Thiero',
        'Austin Reaves',
        'Bronny James',
        'Cameron Carr',
        'Collin Sexton',
        'Dalton Knecht',
        'Jaden Hardy',
        'Jake LaRavia',
        'Jarred Vanderbilt',
        'Kevon Looney',
        'Luka Doncic',
        'Matisse Thybulle',
        'Quentin Grimes',
        'Sandro Mamukelashvili',
        'Walker Kessler',
        'Ziaire Williams'
    ],

    'Memphis Grizzlies': [
        'AJ Johnson',
        'Cam Spencer',
        'Cameron Boozer',
        'Cedric Coward',
        "D'Angelo Russell",
        'G.G. Jackson',
        'Isaiah Stewart',
        'Jaylen Wells',
        'Jerami Grant',
        'Karim Lopez',
        'Kris Murray',
        'Olivier-Maxence Prosper',
        'Quinten Post',
        'Scotty Pippen Jr.',
        'Taj Gibson',
        'Taylor Hendricks',
        'Ty Jerome',
        'Walter Clayton Jr.',
        'Zach Edey'
    ],

    'Miami Heat': [
        'Andrew Wiggins',
        'Bam Adebayo',
        'Bobby Portis',
        'Davion Mitchell',
        'Dru Smith',
        'Giannis Antetokounmpo',
        'J’Vonne Hadley',
        'Klay Thompson',
        'Myron Gardner',
        'Nick Richards',
        'Nikola Jovic',
        'Pelle Larsson',
        'Ryan Conwell',
        'Simone Fontecchio',
        'Tim Hardaway Jr.'
    ],

    'Milwaukee Bucks': [
        'AJ Green',
        'Bogoljub Marković',
        'Brayden Burries',
        'Caris LeVert',
        'Gary Trent Jr.',
        'Jaime Jaquez Jr.',
        'Jericho Sims',
        'Kasparas Jakucionis',
        "Kel'el Ware",
        'Kevin Porter Jr.',
        'Kyle Kuzma',
        'Myles Turner',
        'Nate Ament',
        'Ousmane Dieng',
        'Pete Nance',
        'Ryan Rollins',
        'Tyler Herro'
    ],

    'Minnesota Timberwolves': [
        'Anthony Edwards',
        'Ayo Dosunmu',
        'Cody Williams',
        'Donte DiVincenzo',
        'Isaiah Evans',
        'Jaden McDaniels',
        'Jaylen Clark',
        'Joan Beringer',
        'Jonathan Kuminga',
        'LaMelo Ball',
        "Nah'Shon Hyland",
        'Rudy Gobert',
        'Terrence Shannon Jr.',
        'Trey Lyles'
    ],

    'New Orleans Pelicans': [
        'Bryce McGowens',
        'Caleb Houstan',
        'Christian Koloko',
        'DeAndre Jordan',
        'Dejounte Murray',
        'Derik Queen',
        'Herb Jones',
        'Jeremiah Fears',
        'Jordan Hawkins',
        'Jordan Poole',
        'Karlo Matkovic',
        'Kobe Bufkin',
        'Micah Peavy',
        'Saddiq Bey',
        'Trendon Watford',
        'Trey Murphy III',
        'Yves Missi',
        'Zion Williamson'
    ],

    'New York Knicks': [
        'Andre Drummond',
        'Jalen Brunson',
        'Jordan Clarkson',
        'Jose Alvarado',
        'Josh Hart',
        'Karl-Anthony Towns',
        'Landry Shamet',
        'Mikal Bridges',
        'Miles McBride',
        'Mohamed Diawara',
        'OG Anunoby',
        'Pacome Dadiet',
        'Tyler Kolek'
    ],

    'Oklahoma City Thunder': [
        'Aday Mara',
        'Ajay Mitchell',
        'Alex Caruso',
        'Bennett Stirtz',
        'Cason Wallace',
        'Chet Holmgren',
        'Isaiah Hartenstein',
        'Jalen Williams',
        'Jared McCain',
        'Jaylin Williams',
        'Kenrich Williams',
        'Nikola Topic',
        'Shai Gilgeous-Alexander',
        'Thomas Sorber'
    ],

    'Orlando Magic': [
        'Anthony Black',
        'Desmond Bane',
        'Franz Wagner',
        'Goga Bitadze',
        'J.D. Davison',
        'Jalen Suggs',
        'Jamal Cain',
        'Jase Richardson',
        'Jevon Carter',
        'Jonathan Isaac',
        'Malaki Branham',
        'Nikola Vucevic',
        'Noah Penda',
        'Paolo Banchero',
        'Tristan Da Silva',
        'Wendell Carter Jr.'
    ],

    'Philadelphia 76ers': [
        'Adem Bona',
        'Anfernee Simons',
        'Ariel Hukporti',
        'Dean Wade',
        'Dominick Barlow',
        'Duke Miles',
        'Jabari Walker',
        'Jameer Nelson Jr.',
        'Jaylen Brown',
        'Joel Embiid',
        'Justin Edwards',
        'Kentavious Caldwell-Pope',
        'Labaron Philon Jr.',
        'LeBron James',
        'Saint Thomas',
        'Tacko Fall',
        'Tyrese Maxey',
        'VJ Edgecombe'
    ],

    'Phoenix Suns': [
        'Collin Gillespie',
        'Devin Booker',
        'Dillon Brooks',
        'Haywood Highsmith',
        'Jalen Green',
        'Jamaree Bouyea',
        'Jordan Goodwin',
        'Khaman Maluach',
        'Koa Peat',
        'Luke Kennard',
        'Mark Williams',
        'Miles Bridges',
        'Oso Ighodaro',
        'Rasheer Fleming',
        'Ryan Dunn'
    ],

    'Portland Trail Blazers': [
        'Branden Carlson',
        'Damian Lillard',
        'Deni Avdija',
        'Donovan Clingan',
        'Ja Morant',
        'Jeremy Sochan',
        'Jrue Holiday',
        'Micah Potter',
        'Robert Williams III',
        'Scoot Henderson',
        'Shaedon Sharpe',
        'Sidy Cissoko',
        'Toumani Camara',
        'Vit Krejci',
        'Yang Hansen'
    ],

    'Sacramento Kings': [
        'Alex Karaban',
        'Daeqwon Plowden',
        'Darius Acuff Jr.',
        "De'Andre Hunter",
        'Domantas Sabonis',
        'Dylan Cardwell',
        'Emanuel Sharp',
        'Keegan Murray',
        'Malik Monk',
        'Maxime Raynaud',
        'Nique Clifford',
        'Precious Achiuwa',
        'Zach LaVine'
    ],

    'San Antonio Spurs': [
        'Carter Bryant',
        "De'Aaron Fox",
        'Devin Vassell',
        'Dylan Harper',
        'Harrison Barnes',
        'Jayden Quaintance',
        'Jordan McLaughlin',
        'Julian Champagnie',
        'Keldon Johnson',
        'Luke Kornet',
        'Stephon Castle',
        'Taelon Peter',
        'Tarris Reed Jr.',
        'Tobias Harris',
        'Victor Wembanyama'
    ],

    'Toronto Raptors': [
        'Alijah Martin',
        'Allen Graves',
        'Andre Jackson Jr.',
        'Brandon Ingram',
        'Collin Murray-Boyles',
        'Gradey Dick',
        'Immanuel Quickley',
        "Ja'Kobe Walter",
        'Jakob Poeltl',
        'Jamal Shead',
        'Jamison Battle',
        'Kyle Anderson',
        'Malachi Smith',
        'Nate Bittle',
        'R.J. Barrett',
        'Scottie Barnes',
        'Trayce Jackson-Davis'
    ],

    'Utah Jazz': [
        'Ace Bailey',
        'Brice Sensabaugh',
        'Darryn Peterson',
        'Harrison Ingram',
        'Isaiah Collier',
        'Jaren Jackson Jr.',
        'Jaxson Hayes',
        'Josh Green',
        'Josh Okogie',
        'Jusuf Nurkic',
        'Keyonte George',
        'Kyle Filipowski',
        'Lauri Markkanen',
        'Mohamed Bamba',
        'Sviatoslav Mykhailiuk'
    ],

    'Washington Wizards': [
        'AJ Dybantsa',
        'Alex Sarr',
        'Anthony Davis',
        'Bilal Coulibaly',
        'Bub Carrington',
        'Deandre Ayton',
        'Justin Champagnie',
        'Khris Middleton',
        'Kyshawn George',
        'Trae Young',
        'Tre Johnson',
        'Tre Mann',
        'Tristan Vukcevic',
        'Will Riley'
    ]
}


client = OpenAI(
    api_key="",
    http_client=cast(Any, httpx.Client())
)


BATCH_SIZE = 5

teams = list(league_players.items())

all_results = []

total_batches = (len(teams) + BATCH_SIZE - 1) // BATCH_SIZE


for start_index in range(0, len(teams), BATCH_SIZE):

    batch_items = teams[start_index:start_index + BATCH_SIZE]

    batch = dict(batch_items)

    batch_text = json.dumps(
        batch,
        ensure_ascii=False,
        indent=2
    )

    batch_number = (start_index // BATCH_SIZE) + 1

    print(batch_number)

    response = client.responses.create(
        model="gpt-5.6",
        tools=[
            {
                "type": "web_search"
            }
        ],
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

print(all_results)

client = OpenAI(
    api_key="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",
    http_client = cast(Any, httpx.Client())
)

demands = client.responses.create(
    model="gpt-5.6",
    input=f"""
From this statement, you must parse multiple pieces of information.
The team the user is working for, which players (or what type of players) 
they are looking for to join their team, what players (or what type of players) they are willing to part with for this trade, 
what players (or what type of players) on their team are untouchable, what players you would like avoid acquiring, and what 
weaknesses they want to avoid in acquired players. 

This statement in particular {user_statement}


Return the information in the following format:


“Team Name: [Insert Team Name]
Players Wanted: [Insert Player Names]
Player Types Wanted: [Insert Player Types Wanted]
Players Willing to Part With: [Insert Players Willing to Part With]
Untouchable Players: [Insert Untouchable Players]
Untouchable Player Types: [Insert Untouchable Player Types]
Players to Avoid: [Insert Players to Avoid]
Types of Players to Avoid: [Insert Types of Players to Avoid]
Teams to Avoid: [Insert Teams to Avoid]”

This is the only thing that should be returned. Do not include any other text or explanation.
"""
)

print(demands.output_text)

from typing import Any, cast

import ast
import httpx
from openai import OpenAI


# ============================================================
# INPUTS
# ============================================================

# ============================================================
# OPENAI CLIENT
# ============================================================

client = OpenAI(
    api_key="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",
    http_client=cast(Any, httpx.Client())
)


# ============================================================
# SINGLE API CALL
# ============================================================

response = client.responses.create(
    model="gpt-5.6",
    input=f"""
You are reducing the search space for an NBA trade-generation engine.

You are given:

1. A user's natural-language trade demands.
2. Formatted information for NBA players containing their team,
   Attributes, Strengths, and Weaknesses.

Your job is to reduce the trade search space in TWO stages.

============================================================
USER DEMANDS
============================================================

{demands}

============================================================
FORMATTED PLAYER INFORMATION
============================================================

{all_results}

============================================================
STAGE 1 — REMOVE DISALLOWED PLAYERS AND TEAMS
============================================================

First determine which team the user represents from the USER DEMANDS.

For the USER'S TEAM:

- Remove any player explicitly identified as untouchable.
- Remove any player who matches an untouchable player type, skill,
  attribute, strength, weakness, or characteristic described by the user.
- Use the player's Attributes, Strengths, and Weaknesses to determine
  whether they match the user's restriction.
- Interpret restrictions semantically rather than requiring exact
  keyword matches.

Examples:

If the user says they will not trade good shooters, protect players
whose Strengths indicate things such as:

- shooting
- three-point shooting
- catch-and-shoot ability
- movement shooting
- floor spacing
- pull-up shooting
- perimeter scoring

If the user says they will not trade good defenders, protect players
whose Strengths clearly indicate strong defensive value.

The players remaining on the user's team represent players the user
is willing to trade away.

For EVERY OTHER TEAM:

- Remove any player the user explicitly says they do not want.
- Remove any player matching a player type, weakness, characteristic,
  attribute, or skill that the user says they want to avoid.
- Use each player's Attributes, Strengths, and Weaknesses.
- Interpret the user's wording semantically.
- Do not require exact keyword matches.

Examples:

"bad shooter" may correspond to:

- poor shooting
- weak three-point shooting
- non-shooter
- poor floor spacing
- inefficient perimeter shooting
- unreliable jump shot

"choker" should only match a player if their supplied player information
actually indicates poor clutch performance, playoff decline, choking,
or an equivalent characteristic.

Do NOT infer a characteristic that is not reasonably supported by the
provided player information.

============================================================
TEAM RESTRICTIONS
============================================================

If the user explicitly says they do not want to trade with a team:

- Remove that ENTIRE team.
- Do not include that team in Stage 1.
- Do not include that team in Stage 2.

A team-level restriction takes precedence over a request for an
individual player on that team unless the user's demands clearly
indicate an exception.

============================================================
EXPLICITLY WANTED PLAYERS
============================================================

If the user explicitly names a player they want to acquire:

- Keep that player regardless of player-level avoidance filters.
- Include that player in Stage 1 if their team is still eligible.
- Include that player in Stage 2 if their team is still eligible.

Explicitly wanted players override:

- unwanted player-type filters
- unwanted skill filters
- unwanted weakness filters
- attribute filters
- general player avoidance preferences

However, an explicit TEAM restriction takes precedence.

============================================================
STAGE 1 OUTPUT
============================================================

Stage 1 represents the complete remaining search space.

For the user's team:

Include only players who remain tradeable after applying the user's
untouchable-player and untouchable-player-type restrictions.

For opposing teams:

Include only players who remain allowable after applying:

- explicit player exclusions
- undesirable player-type exclusions
- undesirable skill or attribute exclusions
- team exclusions

Do not remove an opposing player simply because they do not match
something the user wants.

That matching happens in Stage 2.

Stage 1 answers:

"Which players are still allowed to participate in a possible trade?"

============================================================
STAGE 2 — IDENTIFY ACTUAL TRADE TARGETS
============================================================

Using ONLY opposing-team players that survived Stage 1, create a
second and narrower search space.

Do NOT include players from the user's own team in Stage 2.

A player belongs in Stage 2 if ANY of these conditions are true:

1. The user explicitly names that player as someone they want.

OR

2. The player meaningfully matches a player TYPE the user wants.

OR

3. The player meaningfully matches a skill, role, strength,
   characteristic, or attribute the user wants.

Examples include:

- backup center
- starting center
- shooter
- floor spacer
- rim protector
- rebounder
- perimeter defender
- interior defender
- playmaker
- shot creator
- ball handler
- athlete
- young player
- veteran
- high-IQ player
- playoff performer
- clutch player

Interpret these concepts semantically.

For example:

"strong backup center" could reasonably match centers whose supplied
information indicates several relevant traits such as:

- rebounding
- rim protection
- interior defense
- size
- strength
- screening
- efficient finishing
- physicality

The player does NOT need to literally have the phrase
"strong backup center" in their profile.

However:

- Do not stretch weak similarities.
- Do not include someone merely because they vaguely resemble the request.
- There should be meaningful evidence in their supplied Attributes,
  Strengths, or Weaknesses.

Explicitly wanted players must appear in Stage 2 if their team remains
eligible.

Stage 2 answers:

"Which of the allowable opposing players actually match what the user
wants to acquire?"

============================================================
IMPORTANT LOGIC RULES
============================================================

Apply the filters in this order:

1. Determine the user's team.

2. Identify explicitly excluded teams.

3. Identify explicitly untouchable players on the user's team.

4. Identify untouchable player TYPES on the user's team.

5. Identify explicitly unwanted opposing players.

6. Identify unwanted opposing player TYPES.

7. Apply the mandatory exception for explicitly wanted players,
   unless their team is prohibited.

8. Construct Stage 1.

9. From Stage 1 opposing players only, identify players matching
   the user's desired players and desired player types.

10. Construct Stage 2.

Do not invent information.

Use ONLY:

- the user's demands
- the supplied player Attributes
- the supplied player Strengths
- the supplied player Weaknesses
- the supplied team/player associations

Do not use outside knowledge.

Do not perform web searches.

Do not move a player to a different team.

Preserve all player names exactly as supplied.

Preserve all team names exactly as supplied.

Do not include duplicate players.

Do not include teams with zero remaining players.

============================================================
OUTPUT FORMAT
============================================================

Return EXACTLY TWO valid Python dictionaries in raw text.

The FIRST dictionary is Stage 1.

The SECOND dictionary is Stage 2.

Do NOT include:

- headings
- labels
- explanations
- reasoning
- comments
- Markdown
- Markdown code fences
- variable assignments
- introductory text
- concluding text
- any text between the two dictionaries

Do NOT return:

stage_1 = ...

Do NOT return:

stage_2 = ...

Return this structure:

{{
    "Team Name": ["Player Name", "Player Name"],
    "Another Team": ["Player Name", "Player Name"]
}}

{{
    "Team Name": ["Player Name"],
    "Another Team": ["Player Name"]
}}

The response must contain exactly two Python dictionary expressions
and nothing else.
"""
)


# ============================================================
# RAW RESPONSE
# ============================================================

raw_output = response.output_text.strip()

print("\nRAW GPT OUTPUT:\n")
print(raw_output)


# ============================================================
# PARSE THE TWO DICTIONARIES
# ============================================================

def parse_two_dictionaries(text):

    try:
        tree = ast.parse(text, mode="exec")

    except SyntaxError as e:
        raise ValueError(
            "GPT did not return valid Python syntax.\n\n"
            f"Returned output:\n{text}"
        ) from e

    dictionaries = []

    for node in tree.body:

        if isinstance(node, ast.Expr):

            try:
                value = ast.literal_eval(node.value)

            except Exception:
                continue

            if isinstance(value, dict):
                dictionaries.append(value)

    if len(dictionaries) != 2:
        raise ValueError(
            "Expected exactly two Python dictionaries.\n\n"
            f"GPT returned:\n{text}"
        )

    return dictionaries[0], dictionaries[1]


stage_1_players, stage_2_targets = parse_two_dictionaries(
    raw_output
)


# ============================================================
# PRINT PARSED RESULTS
# ============================================================

print("\n" + "=" * 100)
print("STAGE 1 — ALLOWABLE TRADE PLAYERS")
print("=" * 100)

print(stage_1_players)


print("\n" + "=" * 100)
print("STAGE 2 — MATCHING TRADE TARGETS")
print("=" * 100)

print(stage_2_targets)
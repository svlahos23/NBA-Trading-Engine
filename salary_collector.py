import json
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
import openpyxl

TEAM_SLUGS = ["atlanta_hawks", "boston_celtics", "brooklyn_nets", "charlotte_hornets", "chicago_bulls", "cleveland_cavaliers", "dallas_mavericks", "denver_nuggets", "detroit_pistons", "golden_state_warriors", "houston_rockets", "indiana_pacers", "los_angeles_clippers", "los_angeles_lakers", "memphis_grizzlies", "miami_heat", "milwaukee_bucks", "minnesota_timberwolves", "new_orleans_pelicans", "new_york_knicks", "oklahoma_city_thunder", "orlando_magic", "philadelphia_76ers", "phoenix_suns", "portland_trail_blazers", "sacramento_kings", "san_antonio_spurs", "toronto_raptors", "utah_jazz", "washington_wizards"]

def get_nba_contracts():

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36"})
    rows = []
    for i, team_slug in enumerate(TEAM_SLUGS, start = 1):
        url = f"https://www.hoopshype.com/salaries/{team_slug}/"
        print(f"[{i:02d}/30] Pulling {team_slug.replace('_', ' ').title()}...")
        try:
            response = session.get(url, timeout = 30)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            next_data_tag = soup.find("script", id = "__NEXT_DATA__")
            if next_data_tag is None:
                print("No __NEXT_DATA__ found.")
                continue
            next_data = json.loads(next_data_tag.string)
            queries = next_data.get("props", {}).get("pageProps", {}).get("dehydratedState", {}).get("queries", [])
            contracts = None
            for query in queries:
                data = query.get("state", {}).get("data")
                if not isinstance(data, dict):
                    continue
                contract_obj = data.get("contracts")
                if isinstance(contract_obj, dict) and isinstance(contract_obj.get("contracts"), list):
                    contracts = contract_obj["contracts"]
                    break
            if contracts is None:
                print("Contract data not found.")
                continue
            for contract in contracts:
                player = contract.get("player") or {}
                team = player.get("team") or {}
                location = team.get("location")
                nickname = team.get("nickname")
                team_name = f"{location} {nickname}" if location and nickname else nickname or location or team_slug.replace("_", " ").title()
                seasons = contract.get("seasons") or []
                for season in seasons:
                    season_year = season.get("season")
                    try:
                        season_year = int(season_year)
                    except (TypeError, ValueError):
                        continue
                    if season_year < 2026:
                        continue
                    rows.append({"player_id": contract.get("playerID"), "Name": contract.get("playerName"), "team_id": team.get("id"), "team": team_name, "season": f"{season_year}-{str(season_year + 1)[-2:]}", "salary": season.get("salary"), "cap_allocation": season.get("capAllocation"), "team_option": season.get("teamOption"), "player_option": season.get("playerOption"), "two_way": season.get("twoWayContract"), "qualifying_offer": season.get("qualifyingOffer")})
            print(f"    Found {len(contracts)} players.")
        except Exception as e:
            print(f"ERROR: {e}")
        time.sleep(0.25)
    contracts_df = pd.DataFrame(rows)
    if contracts_df.empty:
        raise RuntimeError("No contract data was returned from HoopsHype.")
    contracts_df["salary"] = pd.to_numeric(contracts_df["salary"], errors = "coerce")
    contracts_df["cap_allocation"] = pd.to_numeric(contracts_df["cap_allocation"], errors = "coerce")
    bool_cols = ["team_option", "player_option", "two_way", "qualifying_offer"]
    for col in bool_cols:
        contracts_df[col] = contracts_df[col].fillna(False).astype(bool)
    contracts_df = contracts_df.drop_duplicates(subset = ["player_id", "team_id", "season"]).sort_values(["team", "Name", "season"]).reset_index(drop = True)

    return contracts_df

contracts_df = get_nba_contracts()
print("\nDONE")
print(f"Rows: {len(contracts_df):,}")
print(f"Players: {contracts_df['player_id'].nunique():,}")
print(f"Teams: {contracts_df['team'].nunique():,}")
print("\nSeasons:")
print(contracts_df["season"].drop_duplicates().sort_values().tolist())
print("\nColumns:")
print(contracts_df.columns.tolist())
print("\nSample:")
print(contracts_df.head(30).to_string(index = False))
contracts_df.to_excel("nba_contracts.xlsx", index = False)
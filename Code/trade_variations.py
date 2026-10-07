from itertools import combinations, product


def generate_all_trade_variations(
    wanted_players: dict,
    tradeable_players: dict,
    user_team: str
):
    all_trades = []

    if user_team not in tradeable_players:
        raise ValueError(f"{user_team} not found in tradeable_players.")

    all_teams = [
        team
        for team in tradeable_players
        if team != user_team
    ]

    # ============================================================
    # 2-TEAM TRADES
    # ============================================================

    for other_team in all_teams:

        user_roster = sorted(set(tradeable_players[user_team]))
        other_roster = sorted(set(tradeable_players[other_team]))

        wanted_from_other = set(
            wanted_players.get(other_team, [])
        )

        if not wanted_from_other:
            continue

        # Each user player:
        # 0 = stays
        # 1 = traded to other team
        #
        # Each other-team player:
        # 0 = stays
        # 1 = traded to user

        for user_assignment in product(
            (0, 1),
            repeat=len(user_roster)
        ):
            sent_by_user = tuple(
                player
                for player, move in zip(
                    user_roster,
                    user_assignment
                )
                if move == 1
            )

            if not sent_by_user:
                continue

            for other_assignment in product(
                (0, 1),
                repeat=len(other_roster)
            ):
                sent_to_user = tuple(
                    player
                    for player, move in zip(
                        other_roster,
                        other_assignment
                    )
                    if move == 1
                )

                if not sent_to_user:
                    continue

                # User must actually get at least one wanted player.
                if not any(
                    player in wanted_from_other
                    for player in sent_to_user
                ):
                    continue

                all_trades.append({
                    "trade_type": "2_team",
                    "teams": (
                        user_team,
                        other_team,
                    ),
                    "moves": (
                        {
                            "from_team": user_team,
                            "to_team": other_team,
                            "players": sent_by_user,
                        },
                        {
                            "from_team": other_team,
                            "to_team": user_team,
                            "players": sent_to_user,
                        },
                    ),
                })

    # ============================================================
    # 3-TEAM TRADES
    # ============================================================

    for team_a, team_b in combinations(all_teams, 2):

        teams = (
            user_team,
            team_a,
            team_b,
        )

        rosters = {
            team: sorted(set(tradeable_players[team]))
            for team in teams
        }

        # Flatten all players while retaining original team.
        player_pool = []

        for team in teams:
            for player in rosters[team]:
                player_pool.append(
                    (team, player)
                )

        # For every player:
        #
        # 0 = stays on current team
        # 1 = goes to first possible other team
        # 2 = goes to second possible other team
        #
        # Therefore EVERY allowable player can go to
        # either of the other teams.

        destinations = {}

        for team in teams:
            destinations[team] = [
                destination
                for destination in teams
                if destination != team
            ]

        for assignment in product(
            (0, 1, 2),
            repeat=len(player_pool)
        ):

            moves = {
                (user_team, team_a): [],
                (user_team, team_b): [],
                (team_a, user_team): [],
                (team_a, team_b): [],
                (team_b, user_team): [],
                (team_b, team_a): [],
            }

            for (
                (origin_team, player),
                action
            ) in zip(player_pool, assignment):

                if action == 0:
                    continue

                destination_team = (
                    destinations[origin_team][action - 1]
                )

                moves[
                    (origin_team, destination_team)
                ].append(player)

            # ----------------------------------------------------
            # Remove assignments where nothing happened.
            # ----------------------------------------------------

            nonempty_moves = {
                key: players
                for key, players in moves.items()
                if players
            }

            if not nonempty_moves:
                continue

            # ----------------------------------------------------
            # All THREE teams must actually be involved.
            #
            # A team counts as involved if it sends OR receives
            # at least one player.
            # ----------------------------------------------------

            involved_teams = set()

            for (
                from_team,
                to_team
            ), players in nonempty_moves.items():

                if players:
                    involved_teams.add(from_team)
                    involved_teams.add(to_team)

            if len(involved_teams) != 3:
                continue

            # ----------------------------------------------------
            # User must receive at least one wanted player.
            # ----------------------------------------------------

            received_by_user = []

            for (
                from_team,
                to_team
            ), players in nonempty_moves.items():

                if to_team == user_team:
                    received_by_user.extend(
                        (from_team, player)
                        for player in players
                    )

            gets_wanted_player = any(
                player in set(
                    wanted_players.get(
                        originating_team,
                        []
                    )
                )
                for originating_team, player
                in received_by_user
            )

            if not gets_wanted_player:
                continue

            # ----------------------------------------------------
            # Canonical move representation.
            # Player ordering cannot create duplicates.
            # ----------------------------------------------------

            canonical_moves = []

            for (
                from_team,
                to_team
            ), players in sorted(nonempty_moves.items()):

                canonical_moves.append({
                    "from_team": from_team,
                    "to_team": to_team,
                    "players": tuple(sorted(players)),
                })

            all_trades.append({
                "trade_type": "3_team",
                "teams": teams,
                "moves": tuple(canonical_moves),
            })

    return all_trades



# ================================================================
# GENERATE THE COMPLETE LIST
# ================================================================

all_trade_variations = generate_all_trade_variations(
    wanted_players={'Atlanta Hawks': ['Jock Landale', 'Onyeka Okongwu'], 'Charlotte Hornets': ['Naz Reid'], 'Chicago Bulls': ['Zach Collins'], 'Golden State Warriors': ['Al Horford'], 'Indiana Pacers': ['Jay Huff', 'Larry Nance Jr.'], 'Los Angeles Clippers': ['Brook Lopez'], 'Memphis Grizzlies': ['Isaiah Stewart'], 'Miami Heat': ['Bobby Portis'], 'Oklahoma City Thunder': ['Jaylin Williams'], 'Orlando Magic': ['Goga Bitadze'], 'Portland Trail Blazers': ['Yang Hansen'], 'Utah Jazz': ['Jusuf Nurkic']},
    tradeable_players={'Atlanta Hawks': ['Aaron Wiggins', 'Buddy Hield', 'C.J. McCollum', 'Corey Kispert', 'Henri Veesaar', 'Jock Landale', 'Kingston Flemings', 'Nickeil Alexander-Walker', 'Onyeka Okongwu', 'Ryan Nembhard'], 'Brooklyn Nets': ['Egor Demin', 'Keon Ellis', 'Michael Porter Jr.', 'Mikel Brown Jr.', 'Terance Mann'], 'Charlotte Hornets': ['Brandon Miller', 'Christian Anderson', 'Coby White', 'Grant Williams', 'Grayson Allen', 'Jarkel Joiner', 'Kon Knueppel', 'Liam McNeeley', 'Naz Reid', "Royce O'Neale"], 'Chicago Bulls': ['Jalen Smith', 'Matas Buzelis', 'Norman Powell', 'Patrick Williams', 'Rob Dillingham', 'Tre Jones', 'Zach Collins'], 'Cleveland Cavaliers': ['Donovan Mitchell', 'Mario Hezonja', 'Meleek Thomas', 'Sam Merrill', 'Tyrese Proctor'], 'Dallas Mavericks': ['Kyrie Irving', 'Marcus Sasser', 'Max Christie', 'Morez Johnson Jr.', 'PJ Washington', 'Santi Aldama', 'Sergio de Larrea', 'Tarik Biberovic'], 'Denver Nuggets': ['Aaron Gordon', 'Cameron Johnson', 'DaRon Holmes II', 'Jamal Murray', 'Julian Strawther', 'Nikola Jokic', 'Spencer Jones', 'Trevon Brazile', 'Tyus Jones'], 'Detroit Pistons': ['Cade Cunningham', 'Chaz Lanier', 'Daniss Jenkins', 'Duncan Robinson', 'Gary Harris', 'Isaiah Joe', 'John Collins', 'Taurean Prince'], 'Golden State Warriors': ['Al Horford', 'Brandin Podziemski', "De'Anthony Melton", 'Georges Niang', 'Moses Moody', 'Stephen Curry', 'Will Richard', 'Yaxel Lendeborg'], 'Houston Rockets': ['Bogdan Bogdanovic', 'Bruce Thornton', 'Fred VanVleet', 'Isaiah Crawford', 'Jabari Smith Jr.', 'Kevin Durant', 'Reed Sheppard'], 'Indiana Pacers': ['Aaron Nesmith', 'Andrew Nembhard', 'Jarace Walker', 'Jay Huff', 'Johnny Furphy', 'Larry Nance Jr.', 'Obi Toppin', 'Tyrese Haliburton'], 'Los Angeles Clippers': ['Bradley Beal', 'Brook Lopez', 'Cameron Christie', 'Darius Garland', 'Kawhi Leonard', 'Keaton Wagler', 'Rui Hachimura'], 'Los Angeles Lakers': ['Austin Reaves', 'Bronny James', 'Cameron Carr', 'Collin Sexton', 'Dalton Knecht', 'Jaden Hardy', 'Jake LaRavia', 'Luka Doncic', 'Quentin Grimes', 'Sandro Mamukelashvili'], 'Memphis Grizzlies': ['Cam Spencer', 'Cameron Boozer', "D'Angelo Russell", 'G.G. Jackson', 'Isaiah Stewart', 'Jaylen Wells', 'Jerami Grant', 'Quinten Post', 'Taj Gibson', 'Taylor Hendricks', 'Ty Jerome', 'Walter Clayton Jr.'], 'Miami Heat': ['Andrew Wiggins', 'Bam Adebayo', 'Bobby Portis', 'Davion Mitchell', 'Dru Smith', 'Klay Thompson', 'Myron Gardner', 'Nikola Jovic', 'Ryan Conwell', 'Simone Fontecchio', 'Tim Hardaway Jr.'], 'Milwaukee Bucks': ['AJ Green', 'Bogoljub Marković', 'Brayden Burries', 'Gary Trent Jr.', 'Kasparas Jakucionis', "Kel'el Ware", 'Kevin Porter Jr.', 'Myles Turner', 'Pete Nance', 'Ryan Rollins', 'Tyler Herro'], 'Minnesota Timberwolves': ['Anthony Edwards', 'Donte DiVincenzo', 'Isaiah Evans', 'Jaden McDaniels', 'LaMelo Ball', "Nah'Shon Hyland", 'Rudy Gobert', 'Trey Lyles'], 'New Orleans Pelicans': ['Bryce McGowens', 'Caleb Houstan', 'Jordan Hawkins', 'Jordan Poole', 'Karlo Matkovic', 'Micah Peavy', 'Saddiq Bey', 'Trey Murphy III'], 'New York Knicks': ['Andre Drummond', 'Jose Alvarado', 'Josh Hart', 'Mohamed Diawara'], 'Oklahoma City Thunder': ['Alex Caruso', 'Bennett Stirtz', 'Chet Holmgren', 'Jared McCain', 'Jaylin Williams', 'Kenrich Williams', 'Shai Gilgeous-Alexander'], 'Orlando Magic': ['Desmond Bane', 'Jase Richardson', 'Jevon Carter', 'Jonathan Isaac', 'Nikola Vucevic', 'Tristan Da Silva'], 'Philadelphia 76ers': ['Anfernee Simons', 'Dean Wade', 'Joel Embiid', 'Justin Edwards', 'Kentavious Caldwell-Pope', 'Labaron Philon Jr.', 'Tacko Fall', 'Tyrese Maxey'], 'Phoenix Suns': ['Collin Gillespie', 'Dillon Brooks', 'Haywood Highsmith', 'Luke Kennard', 'Rasheer Fleming'], 'Portland Trail Blazers': ['Branden Carlson', 'Damian Lillard', 'Jrue Holiday', 'Micah Potter', 'Vit Krejci', 'Yang Hansen'], 'Sacramento Kings': ['Alex Karaban', 'Darius Acuff Jr.', 'Emanuel Sharp', 'Malik Monk', 'Zach LaVine'], 'San Antonio Spurs': ['Devin Vassell', 'Harrison Barnes', 'Jordan McLaughlin', 'Julian Champagnie', 'Keldon Johnson', 'Taelon Peter', 'Tobias Harris', 'Victor Wembanyama'], 'Toronto Raptors': ['Allen Graves', 'Brandon Ingram', 'Gradey Dick', 'Immanuel Quickley', 'Jamison Battle'], 'Utah Jazz': ['Brice Sensabaugh', 'Darryn Peterson', 'Harrison Ingram', 'Jaren Jackson Jr.', 'Josh Green', 'Jusuf Nurkic', 'Keyonte George', 'Lauri Markkanen', 'Mohamed Bamba', 'Sviatoslav Mykhailiuk'], 'Washington Wizards': ['Bub Carrington', 'Justin Champagnie', 'Khris Middleton', 'Kyshawn George', 'Trae Young', 'Tre Johnson', 'Tre Mann', 'Tristan Vukcevic']},

    user_team="New York Knicks"
)

print(f"Total trade variations: {len(all_trade_variations):,}")
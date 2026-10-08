# NBA Trading Engine

An ongoing project to build an automated NBA trade engine capable of identifying potential trades and determining whether those trades are legal under the NBA Collective Bargaining Agreement (CBA).

The project combines NBA contract data with CBA trade restrictions so that potential transactions can eventually be generated, filtered, and evaluated automatically.

## Repository Files

### `Code/salary_collector.py`

Python script responsible for collecting NBA player salary and contract information.

The script gathers contract data for players across NBA teams and organizes the results into a structured dataset that can be used by the trading engine.

The collected data is stored in:

`nba_trade_engine_data.xlsx`

This allows the underlying salary information to be updated without manually maintaining the contract database.

---

### `Code/user_preferences.py`

Python script responsible for interpreting a user's natural-language trade preferences and reducing the NBA-wide player pool into a smaller search space for trade generation.

The script uses GPT to evaluate current NBA players based on attributes, strengths, and weaknesses, then parses the user's trade request to identify information such as:

* The team the user represents
* Players and player types the user wants to acquire
* Players the user is willing to trade
* Untouchable players and player types
* Players and player types the user wants to avoid
* Teams the user does not want to trade with

The script then performs two stages of filtering.

Stage 1 produces the complete set of players who are still allowed to participate in a possible trade after applying the user's restrictions.

Stage 2 identifies the opposing players from that remaining search space who actually match the user's desired players, roles, skills, attributes, or player types.

These outputs provide the `tradeable_players` and `wanted_players` inputs used by the trade-generation portion of the engine.

---

### `Code/trade_variations.py`

Python script responsible for generating possible player-trade combinations from the reduced player pools produced by the user-preference stage.

The script accepts:

* The user's team
* Players who are allowed to participate in trades
* Players the user wants to acquire

It generates both two-team and three-team trade variations.

For two-team trades, the script evaluates possible combinations of outgoing and incoming players while requiring the user's team to receive at least one wanted player.

For three-team trades, players can be routed between any of the three participating teams. The script ensures that all three teams are actually involved in the transaction and that the user's team receives at least one wanted player.

The resulting trades are stored in a standardized structure containing:

* Trade type
* Participating teams
* Origin team for each group of players
* Destination team for each group of players
* Players included in each movement

These generated trade variations can then be passed to the CBA-validation layer to remove transactions that are not legally permissible.

---

### `Code/cba_trade_rules.py`

Python module responsible for validating generated trade variations against NBA Collective Bargaining Agreement restrictions.

The module is designed to operate on the general trade structure produced by `trade_variations.py` rather than on any specific team, player request, or individual generated result.

It loads financial and contract information from:

`nba_trade_engine_data.xlsx`

The validator evaluates each participating team independently and checks applicable trade restrictions including:

* Salary matching
* Use of available cap room
* Standard Traded Player Exception salary matching
* Aggregated salary matching
* Expanded Traded Player Exception rules
* First Apron restrictions
* Second Apron restrictions
* Existing trade exceptions
* Player trade eligibility restrictions
* Recently signed player restrictions
* Recently acquired player aggregation restrictions
* Sign-and-trade restrictions
* Base Year Compensation considerations
* Rookie-extension poison-pill treatment
* Non-guaranteed salary treatment
* Trade bonuses
* Trade-deadline restrictions
* Cash considerations
* Draft-pick restrictions
* Stepien Rule considerations
* Multi-team trade validation

The module filters the generated trade candidates and retains the transactions that do not violate the CBA rules supported by the available data.

Additional player- or team-specific information that is not available in the underlying workbook can be supplied through overrides so that restrictions requiring information such as signing dates, acquisition dates, consent rights, or special trade salary treatment can also be evaluated.

The module also supports rejection diagnostics when information about why a proposed trade failed CBA validation is needed.

---

### `Code/nba_two_team_trade_engine.py`

Python module that combines two-team trade generation and NBA Collective Bargaining Agreement validation into a single, optimized process.

The module integrates the functionality of `trade_variations.py` and `cba_trade_rules.py`, allowing possible trades to be generated and evaluated against applicable CBA restrictions without first constructing an exhaustive list of every player combination.

The function accepts only two inputs:

* The file path to `nba_trade_engine_data.xlsx`
* A structured user-preferences string containing the user's team, Stage 1 tradeable players, and Stage 2 wanted players

The module uses the filtered player pools produced by `user_preferences.py` to identify eligible trading partners and construct possible two-team transactions.

To improve computational efficiency, the engine:

* Generates only two-team trades involving the user's team
* Excludes teams and players removed during the user-preference filtering stage
* Requires both participating teams to send and receive at least one player
* Requires the user's team to acquire at least one wanted player
* Avoids generating duplicate player combinations
* Applies salary-based pruning to eliminate financially invalid combinations before full trade validation
* Uses the existing CBA validation logic to evaluate remaining candidates
* Returns trades incrementally rather than storing every possible combination in memory

Each generated trade is evaluated against applicable NBA CBA restrictions, including salary matching, salary cap considerations, first and second apron restrictions, player trade eligibility, and other contract-specific limitations supported by the underlying data.

The resulting trades retain a standardized structure identifying the participating teams, players exchanged, and direction of each player movement.

Because certain CBA restrictions require information not consistently available in the workbook, such as exact signing dates, acquisition dates, consent rights, and special salary treatment, the engine distinguishes between trades with no detected violations and transactions requiring additional verification.

The purpose of this module is to reduce the computational cost of exhaustive trade generation while preserving all valid two-team player combinations within the user-defined search space and the restrictions that can be evaluated from the available data.

---

### `Data/nba_trade_engine_data.xlsx`

The current NBA player contract and salary dataset generated by `salary_collector.py`.

This file provides the financial information needed to evaluate trades, including player salaries and contract details.

The trading engine will use this information when determining:

* Salary matching requirements
* Team salary implications
* Trade eligibility
* CBA-related financial restrictions
* Whether a proposed player combination can legally be traded

---

### `Text Files/NBA Trade CBA Rules Reference.txt`

A reference document containing the NBA Collective Bargaining Agreement rules that affect trades.

The purpose of this file is to identify the restrictions that the trading engine will eventually need to enforce programmatically.

These rules include areas such as:

* Salary matching
* Salary cap and tax-apron restrictions
* First and second apron trade limitations
* Aggregation restrictions
* Sign-and-trade rules
* Recently signed player restrictions
* Recently traded player restrictions
* Base Year Compensation
* Trade exceptions
* Draft-pick restrictions
* Stepien Rule considerations
* Cash and other transaction limitations

This file currently serves as the rule specification for the trade-validation portion of the project.

---

### `Text Files/ChatGPT Inputs.txt`

A reference document containing the structured prompts that will allow users to communicate with the trading engine using natural language rather than predetermined keywords or inputs.

The purpose of this file is to make the engine more dynamic by connecting it to a GPT API that can interpret what a user is asking for, identify the team they represent, and convert their requests into structured information that the trading engine can use.

The GPT API will be used to:

* Parse natural-language trade requests into structured trade preferences
* Identify players and player types a user wants to acquire
* Identify players and player types a user is willing to trade
* Identify untouchable players and player types
* Identify players, player types, and teams the user wants to avoid
* Generate attributes, strengths, and weaknesses for players
* Match player characteristics to the user's requested trade criteria
* Reduce the number of tradeable players based on the user's requirements
* Reduce the overall trade search space before possible trades are generated

Reducing the search space is important because of the extremely large number of possible player combinations across the NBA. By filtering players according to the user's objectives and player characteristics before trade generation begins, the engine can focus on trades that the user might realistically want to pursue.

---

## Development Roadmap

The project is being developed incrementally.

### Current

* Collect NBA player contract and salary data
* Create a structured contract dataset
* Research and document NBA CBA trade restrictions

### Longer Term

The goal is to move beyond simply checking whether a manually entered trade works.

The engine will eventually be able to search possible transactions automatically and return trades that satisfy:

1. NBA CBA requirements
2. Player and draft-pick eligibility rules
3. Team salary constraints
4. User-defined trade objectives

## Project Status

This repository is under active development. The current focus is building the data and rule infrastructure required before automated trade generation can be implemented.

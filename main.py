import pandas as pd
import pyrankvote
from pyrankvote import Candidate, Ballot


# Function to process elections from a specified sheet
def process_election(sheet_name, number_of_seats):
    # Read the Excel file assuming the specified sheet contains the header in the first row
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    #print(f"DataFrame Loaded for {sheet_name}:", df.head())

    # Ensure there are no NaN values in the headers, if there are, replace them
    df.columns = [x if pd.notna(x) else f'Unnamed: {i}' for i, x in enumerate(df.columns)]
    #print(f"Corrected Headers for {sheet_name}:", df.columns)

    # Create Candidate objects from the column names
    candidates = [Candidate(name) for name in df.columns]
    #print(f"Candidates List for {sheet_name}:", candidates)

    # Create a list of Ballot objects from the DataFrame
    raw_ballots = []
    for index, row in df.iterrows():
        # Build the ranked list of Candidate objects, ignoring NaN values
        ranked_candidates = [candidates[df.columns.get_loc(name)] for name in row if
                             pd.notna(name) and name in df.columns]
        if ranked_candidates:
            ballot = Ballot(ranked_candidates)
            raw_ballots.append(ballot)
        #print(f"Created Ballot for {sheet_name}:", ballot)

    # Convert Ballots to indices
    candidate_index_map = {candidate.name: i for i, candidate in enumerate(candidates)}
    print(f"Candidate Index Map for {sheet_name}:", candidate_index_map)
    ballots_indices = [[candidate_index_map[candidate.name] for candidate in ballot.ranked_candidates] for ballot in
                       raw_ballots]
    print(f"Ballot Indices for {sheet_name}:", ballots_indices)

    # Run the STV election
    election_result = pyrankvote.single_transferable_vote(candidates, raw_ballots, number_of_seats=number_of_seats)

    # Get the winners
    winners = election_result.get_winners()

    # Print the election result and the winners
    print(f"Election Results for {sheet_name}")
    print(election_result)
    print(f"The winners are: {[winner.name for winner in winners]}")
    print("############################################")

    return winners


# Path to the Excel file
file_path = 'voting_data.xlsx'  # Update this to your file path

# Process elections for President
process_election('President', number_of_seats=2)

# Process elections for Advisor from 'Advisor'
process_election('Advisor', number_of_seats=2)

# Process elections for Outreach Officer from 'Outreach'
process_election('Outreach', number_of_seats=1)

# Process elections for Media Officer
process_election('Media', number_of_seats=1)

# Process elections for Secretary
process_election('Secretary', number_of_seats=1)

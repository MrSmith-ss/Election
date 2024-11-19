import pandas as pd
import openpyxl
import os
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import streamlit as st

@st.cache_data(ttl=36000)
def load_data(sheet_name):
    # Define the Excel file path relative to the script's location
    # Excel data generated from:  #https://www.fec.gov/introduction-campaign-finance/election-results-and-voting-information/
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_excel_file = os.path.join(script_dir, "2000-2024 Election Data.xlsx")
    # Load the data once
    return pd.read_excel(input_excel_file, sheet_name='Output')

@st.cache_data
def filter_data(df, start_year, end_year):
    # Filter data based on the selected year range
    return df[(df['Year'].astype(int) >= start_year) & 
              (df['Year'].astype(int) <= end_year)]

@st.cache_data
def sort_states_by_mode(filtered_df, states, mode):
    if mode in ['D', 'R']:
        differences = []
        for state in states:
            state_df = filtered_df[filtered_df['State'] == state]
            target_party = 'Democrat' if mode == 'D' else 'Republican'
            
            # Get the 2020 vote count for the target party
            votes_2020 = state_df[state_df['Year'] == 2020][target_party].values
            if votes_2020.size > 0:
                non_2020_max = state_df[state_df['Year'] != 2020][target_party].max()
                if pd.notna(non_2020_max):
                    difference = votes_2020[0] - non_2020_max
                    differences.append((state, difference))
                else:
                    differences.append((state, 0))
            else:
                differences.append((state, 0))
        
        # Sort states based on the computed differences
        return [state for state, _ in sorted(differences, key=lambda x: x[1], reverse=True)]
    else:
        return states

def generate_all_states_chart(df, start_year, end_year, parties, mode='A', selected_state=None):
    # Filter data for the selected year range
    filtered_df = filter_data(df, start_year, end_year)

    # Get the unique states from the data
    unique_states = st.session_state.setdefault('unique_states', df['State'].unique())

    sorted_states = list(sort_states_by_mode(filtered_df, unique_states, mode))

    







    # Store current filter parameters as a list
    current_filter_params = [start_year, end_year, mode, sorted(parties)]  # sorted(parties) to maintain order

    # Retrieve previous filter parameters (if available)
    previous_filter_params = st.session_state.setdefault('filter_params2', [])

    # Retrieve old state (if available)
    old_state = st.session_state.setdefault('old_state', sorted_states[0])

    # Retrieve old state (if available)
    flag2 = st.session_state.setdefault('flag2', 0)

    # Compare the current filter params to the previous ones
    if current_filter_params != previous_filter_params:
        print("Filters have changed.")
        flag = 2
        flag2 = 1
         # Retrieve the selected state from session state or use the first state
        selected_state = old_state
    else:
        print("Filters have not changed.")
        flag = 1
        if flag2 == 1:
            selected_state = old_state
            flag2 = 0
        
    # Save the current filter parameters in session_state for later comparison
    st.session_state['filter_params2'] = current_filter_params

    #print(f"Saving flag2 to: {flag2}")
    st.session_state['flag2'] = flag2
   
    # Save the selected_state
    if flag == 2:
        print(f"Saving old_state EARLY to: {selected_state}")
        st.session_state['old_state'] = selected_state


    # Check if the selected_state is valid
    if selected_state not in sorted_states:
        print(f"Selected state {selected_state} is not in the list of available states. Resetting to {sorted_states[0]}.")
        selected_state = sorted_states[0]

    # Streamlit sidebar to select a state (scrollable list box)
    print(f"Eary Selected State: {selected_state}")
    selected_state_index = sorted_states.index(selected_state)
    print(f"Selected State Index: {selected_state_index}")
    selected_state = st.sidebar.radio("Select a State", sorted_states, index=selected_state_index)
    print(f"Selected State Late: {selected_state}")
    # Save the selected_state
    print(f"Late Selected State: {selected_state}")
    if flag == 1:
        print(f"Saving old_state LATER to: {selected_state}")
        st.session_state['old_state'] = selected_state



    










    # Set up the figure for the plot
    fig, ax = plt.subplots(figsize=(10, 6),dpi=120)

    # Filter data for the selected state
    state_df = filtered_df[filtered_df['State'] == selected_state]

    # Define custom colors for the parties
    party_colors = {
        'Republican': 'red',
        'Democrat': 'blue',
        'Other': 'yellow'
    }

    # Plotting the bar chart for selected parties with custom colors
    state_df.set_index('Year')[parties].plot(kind='bar', ax=ax, width=0.8, color=[party_colors.get(party, 'gray') for party in parties])

    # Determine the target party for overvote if in "D" or "R" mode
    target_party = None
    difference = None
    if mode == 'D':
        target_party = 'Democrat'
    elif mode == 'R':
        target_party = 'Republican'

    # If a target party is set, calculate the difference and add the lines
    if target_party:
        # Get the 2020 vote count for the target party
        votes_2020 = state_df[state_df['Year'] == 2020][target_party].values
        if votes_2020.size > 0:
            # Calculate the highest vote count from non-2020 years
            non_2020_max = state_df[state_df['Year'] != 2020][target_party].max()
            if pd.notna(non_2020_max):
                difference = votes_2020[0] - non_2020_max
                ax.axhline(y=votes_2020[0], color='purple', linestyle='--', linewidth=2, label=f'{target_party} 2020')
                ax.axhline(y=non_2020_max, color='gray', linestyle='--', linewidth=2, label=f'{target_party} (Highest non-2020)')

    
    # Dictionary mapping state abbreviations to full names
    state_abbr_to_full = {
        "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
        "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "FL": "Florida", "GA": "Georgia",
        "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
        "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
        "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri",
        "MT": "Montana", "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey",
        "NM": "New Mexico", "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
        "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
        "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont",
        "VA": "Virginia", "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming"
    }

    # Look up the full state name from the abbreviation
    state_full_name = state_abbr_to_full.get(selected_state, selected_state)

    # Set the title, labels, and legend
    ax.set_title(f"Votes in {state_full_name} ({start_year}-{end_year})", fontsize=16) # Set the chart title with the full state name
    ax.set_xlabel('Year', fontsize=12)
    ax.set_ylabel('Votes', fontsize=12)
    ax.set_xticklabels([str(x) for x in state_df['Year']], rotation=45, ha='right')

    # Add the calculated difference to the legend with commas if available
    if difference is not None:
        formatted_difference = f"{difference:,.0f}"  # Format with commas
        legend_title = f'{target_party} 2020 Overvote: {formatted_difference}'
        
        # Set the color of the legend title text based on the party
        legend_color = 'blue' if target_party == 'Democrat' else 'red'
        
        # Customizing the legend title color using font properties
        legend = ax.legend(title=legend_title, loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=3)
        
        # Change the font properties for the legend title
        legend.get_title().set_fontsize(13)
        legend.get_title().set_fontweight('bold')
        legend.get_title().set_color(legend_color)
    else:
        ax.legend(title='Party', loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=3)

    # Display the plot in Streamlit

    st.pyplot(fig)



def main():
    # Inject custom CSS
    st.markdown(
        """
        <style>
        /* Apply body margin and padding reset */
        body {
            margin: 0;
            padding: 0;
        }

        /* Adjust the block-container size */
        .block-container {
            max-width: 70% !important;
            margin: auto;
            padding: 0;
            overflow: hidden;  /* Prevent scrolling */
        }

        /* Prevent the chart figure from introducing unwanted space */
        .stPyplot > div {
            padding: 0;  /* Remove padding around the chart */
        }
        </style>
        """, 
        unsafe_allow_html=True
    )
    
     # Load the data from the Output sheet
    df = load_data("Output")
    
    # Sidebar inputs with 4-year steps for the start and end years
    start_year = st.sidebar.number_input("Start Year", min_value=2000, max_value=2024, value=2000, step=4)
    end_year = st.sidebar.number_input("End Year", min_value=2000, max_value=2024, value=2024, step=4)
    
    # Static checkboxes for party selection
    st.sidebar.subheader("Select Parties")
    republican = st.sidebar.checkbox("Republican", value=True)
    democrat = st.sidebar.checkbox("Democrat", value=True)
    other = st.sidebar.checkbox("Other", value=True)
    
    # Create the parties list based on checkbox selections
    parties = []
    if republican:
        parties.append("Republican")
    if democrat:
        parties.append("Democrat")
    if other:
        parties.append("Other")

    # Update the mode names in the selectbox
    mode = st.sidebar.selectbox(
        "Mode",
        ["No Filter", "2020 Republican Overvote", "2020 Democrat Overvote"]
    )

    # Map the mode names to the original values
    mode_mapping = {
        "No Filter": "A",
        "2020 Republican Overvote": "R",
        "2020 Democrat Overvote": "D"
    }

    # Display disclaimer if the end year is 2024
    if end_year == 2024:
        st.markdown(
            "<span style='color: red; font-weight: bold;'>&#42;&#42;&#42;2024 Election Data Is Preliminary&#42;&#42;&#42;</span>",
            unsafe_allow_html=True
        )

    # Generate the chart with the mapped mode value and the selected state
    generate_all_states_chart(df, start_year, end_year, parties, mode_mapping[mode])

if __name__ == "__main__":
    main()
import datetime
import streamlit as st
import pandas as pd

# Load the Excel file only once at the beginning
xl_file = 'data.xlsx'
dfs = pd.read_excel(xl_file, sheet_name=None)


# Function to display the selected sheet and its records
def display_sheet(sheet_name):
    df = st.session_state.dfs[sheet_name]
    st.write(df)


# Function to add a new record to the selected sheet
def add_record(sheet_name, data):
    df = st.session_state.dfs[sheet_name]
    max_id = df['ID'].max() if not df.empty else 0
    new_id = max_id + 1
    data['ID'] = new_id
    new_row = pd.DataFrame(data, index=[0])
    df = pd.concat([df, new_row], ignore_index=True)
    with pd.ExcelWriter(xl_file) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    st.session_state.dfs[sheet_name] = df
    st.success("Record added successfully!")


# Function to update a record in the selected sheet
def update_record(sheet_name, record_id, data):
    df = st.session_state.dfs[sheet_name]
    if record_id in df['ID'].values:
        df.loc[df['ID'] == record_id, data.keys()] = list(data.values())
        with pd.ExcelWriter(xl_file) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        st.session_state.dfs[sheet_name] = df
        st.success("Record updated successfully!")
    else:
        st.warning("Record with the specified ID does not exist.")


# Function to delete a record from the selected sheet based on ID
def delete_record(sheet_name, record_id):
    df = st.session_state.dfs[sheet_name]
    if record_id in df['ID'].values:
        df.drop(df[df['ID'] == record_id].index, inplace=True)
        # Decrement ID for records with IDs greater than the deleted ID
        df.loc[df['ID'] > record_id, 'ID'] -= 1
        with pd.ExcelWriter(xl_file) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        st.session_state.dfs[sheet_name] = df
        st.success("Record deleted successfully!")
    else:
        st.warning("Record with the specified ID does not exist.")


# Main Streamlit app
def main():
    st.title("Excel Data Management App")

    # Initialize session state to store dataframes for each sheet
    if 'dfs' not in st.session_state:
        st.session_state.dfs = dfs

    sheet_choice = st.sidebar.selectbox("Select Sheet", list(st.session_state.dfs.keys()))

    action = st.sidebar.radio("Select Action", ("View", "Add", "Update", "Delete"))

    if action == "View":
        display_sheet(sheet_choice)

    elif action == "Add":
        st.subheader("Add Record")
        new_data = {}
        required_fields = ['ID', 'DATE ENTREE', 'HEURE ENTREE', 'EXPEDITEUR', 'DESTINATAIRE', 'OBJET', 'DATE SORTIE',
                           'HEURE SORTIE']
        for col in st.session_state.dfs[sheet_choice].columns:
            if col == 'ANNOTATIONS':  # Handle the annotations field separately
                new_data[col] = st.text_input(col, key=f"{col}_{sheet_choice}").title()
            elif col in ['DATE ENTREE', 'DATE SORTIE']:  # Handle date inputs
                new_data[col] = st.date_input(f"{col}*", key=f"{col}_{sheet_choice}")
            elif col in ['HEURE ENTREE', 'HEURE SORTIE']:  # Handle time inputs
                new_data[col] = st.time_input(f"{col}*", key=f"{col}_{sheet_choice}")
            elif col != 'ID':
                new_data[col] = st.text_input(f"{col}*", key=f"{col}_{sheet_choice}").title()
        if st.button("Add Record"):
            # Check if any required field is empty
            if any(new_data[col] == '' for col in required_fields if col in new_data):
                st.error("Please fill in all required fields.")
            else:
                # Ensure that date inputs are converted to datetime.date objects
                for col, value in new_data.items():
                    if isinstance(value, datetime.date):
                        new_data[col] = value.strftime("%Y-%m-%d")  # Convert to string if needed
                add_record(sheet_choice, new_data)

    elif action == "Update":
        st.subheader("Update Record")
        record_id = st.number_input("Enter ID of Record to Update", min_value=1,
                                    value=1, step=1)
        record_to_update = st.session_state.dfs[sheet_choice][st.session_state.dfs[sheet_choice]['ID'] == record_id]
        if not record_to_update.empty:
            updated_data = {}
            for col in st.session_state.dfs[sheet_choice].columns:
                if col == 'DATE ENTREE':
                    updated_data[col] = st.date_input(col, value=pd.to_datetime(record_to_update[col].values[0]),
                                                      key=f"{col}_{sheet_choice}")
                elif col == 'DATE SORTIE':
                    updated_data[col] = st.date_input(col, value=pd.to_datetime(record_to_update[col].values[0]),
                                                      key=f"{col}_{sheet_choice}")
                elif col == 'HEURE ENTREE':
                    updated_data[col] = st.time_input(col, value=record_to_update[col].values[0],
                                                      key=f"{col}_{sheet_choice}")
                elif col == 'HEURE SORTIE':
                    updated_data[col] = st.time_input(col, value=record_to_update[col].values[0],
                                                      key=f"{col}_{sheet_choice}")

                elif col != 'ID':
                    updated_data[col] = st.text_input(col, value=record_to_update[col].values[0]).title()
                update_record(sheet_choice, record_id, updated_data)
        else:
            st.warning("Record with the specified ID does not exist.")

    elif action == "Delete":
        st.subheader("Delete Record")
        record_id = st.number_input("Enter ID of Record to Delete", min_value=1,
                                    value=1, step=1)
        if st.button("Delete Record"):
            delete_record(sheet_choice, record_id)


if __name__ == "__main__":
    main()

# Streamlit Excel Data Management App

This Streamlit app allows users to view, add, update, and delete records in an Excel file. It provides a user-friendly interface to interact with the data stored in different sheets of the Excel file.

## Features

- **View**: Users can select a sheet from the Excel file and view all records in that sheet.
- **Add**: Users can add new records to the selected sheet by providing necessary information through the app interface.
- **Update**: Users can update existing records in the selected sheet by specifying the record ID and modifying the required fields.
- **Delete**: Users can delete a record from the selected sheet by specifying the record ID.

## Usage

1. Clone the repository to your local machine.
2. Install the required dependencies using `pip install -r requirements.txt`.
3. Place your Excel file (`data.xlsx`) in the root directory of the project.
4. Run the Streamlit app using `streamlit run app.py`.
5. Interact with the app through the browser interface to manage your Excel data.

## Dependencies

- `pandas`: For data manipulation and Excel file handling.
- `streamlit`: For building the web application interface.
- `datetime`: For handling date and time inputs.

## File Structure

- `app.py`: Main Python script containing the Streamlit application code.
- `data.xlsx`: Excel file containing the data to be managed.
- `requirements.txt`: List of Python dependencies required to run the application.

## Contributing

Contributions are welcome! If you have any suggestions, bug reports, or feature requests, please open an issue or submit a pull request.



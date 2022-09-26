from utils import create_authorized_service, update_values, format_cells, create_pivot_tables
from data_loader import get_data

if __name__ == "__main__":
    
    # create a sheets service
    service = create_authorized_service()

    # Your spreadsheet id
    spreadsheet_id = "YOUR_ID"

    # get the data from csv file
    data = get_data("expenses.csv")

    # insert the data in the google sheets 
    update_values(service, spreadsheet_id, data)

    # format the cells of the google sheets 
    format_cells(service, spreadsheet_id)

    # create pivot tables
    create_pivot_tables(service, spreadsheet_id, data)
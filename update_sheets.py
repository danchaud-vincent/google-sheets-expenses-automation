from utils import create_authorized_service, update_values
from data_loader import get_data

if __name__ == "__main__":
    
    # create a sheets service
    service = create_authorized_service()

    # Your spreadsheet id
    sheet_id = "1ccfchG81UKwMHwjiuxPuN2ffISoijB3HR4ryl03XP_k"

    # get the data from csv file
    data = get_data("expenses.csv")

    # insert the data in the google sheets 
    update_values(service, sheet_id, data)

    # format the cells of the google sheets 

    # create pivot tables
from utils import create_authorized_service, create_sheet

if __name__ == "__main__":
    
    # create google sheets service
    service = create_authorized_service()

    # create a new google sheet
    create_sheet(service)



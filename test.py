from datetime import datetime, timedelta

def get_confirmation():
    choice = input("Do you want to create yesterday's report? (Press 'N' or 'n' for No, any other key for Yes): ").strip()
    return choice.lower() != 'n'

# Example usage
if get_confirmation():
    yesterday = datetime.now() - timedelta(days=1)
    yesterday = yesterday.strftime("%d-%b-%y")  # Example: 10-Mar-25
    print(yesterday)
else:
    date = int(input("Enter how many days ago from today's date you want to get: "))
    date = datetime.now() - timedelta(days=date)
    date = date.strftime("%d-%b-%y")
    print(date)
    print("You chose NO!")

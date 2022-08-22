from excel_worker import *

def data_info():
    nrow, ncol = list(map(int, input("Enter the number of rows and columns: ").split()))
    print()
    col_names = list(map(str, (input("Enter the name of column {} ".format(i)) for i in range(1, ncol + 1))))

    print("\nEnter only string/integer/float")
    col_features = list(map(str, (input("What date type u need for {} column? ".format(name)).lower() for name in col_names)))
    noise = True if (input("\nDo u need NULL value in dataset? yes/no ").lower()) == "yes" else False
    noise_lvl = int(input("Enter the noise percent of noise you need 0-100 ")) if noise else 10
    sl_no = True if input("Do you want to add serial number? yes/no ").lower() == 'yes' else False

    if sl_no:
        adding_serial_number(nrow)

    file_name = input("Enter the file name you want to save the dataset as? ")

    # create the file with the details

    creating_excel(nrow, ncol, col_names, col_features, noise, noise_lvl, sl_no, file_name)


if __name__ == "__main__":
    data_info()

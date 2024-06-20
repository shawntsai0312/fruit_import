from categorizeWithFile import categorize
from toCSV import toCSV
from toXLSX import toXLSX

def main():
    date = input("Enter the date (in MMDD format): ")
    cherrys = categorize(date)
    toCSV(cherrys)
    toXLSX(cherrys, output_file=f'{date}.xlsx')

if __name__ == "__main__":
    main()

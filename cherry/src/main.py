from categorize import categorize
from toCSV import toCSV
from toXLSX import toXLSX

def main():
    date = input("Enter the date (in MMDD format): ")
    cherrys = categorize(date)
    # toCSV(cherrys)
    toXLSX(cherrys)

if __name__ == "__main__":
    main()

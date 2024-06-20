import re
from datetime import datetime
import combineRows


def transform_date(date_str):
    # Parse the input string into a datetime object
    date_obj = datetime.strptime(date_str, '%d-%b')

    # Format the datetime object into the desired string format
    transformed_date = date_obj.strftime('%m%d')

    return transformed_date


def find_element_index(lst, element):
    try:
        return lst.index(element)
    except ValueError:
        return -1  # Return -1 if the element is not found


def categorize(data):

    match = re.search(r'Customer Name(.*?)Travis Hallman', data, re.DOTALL)

    # 如果找到了匹配，那麼就取得 "Customer Name" 之後的所有文字
    if match:
        content = match.group(1)

    content = content.replace("CTNS\n\n", "CTNS\n")
    content = content.replace("SHINE FOOD CORPORATION\n", "")
    content = content.replace("Region  Code\n", "")
    content = content.replace("RegionÂ  Code\n", "")
    content = content.replace("TA\n", "")
    content = content[1:]

    # print(content)

    # Converting the string data into a list of lines
    lines = content.strip().split('\n')

    arrivalDateIndex = find_element_index(lines, 'Arrival Date')
    varietyCodeIndex = find_element_index(lines, 'Variety Code')
    packCodeIndex = find_element_index(lines, 'Pack Code')
    sizeCodeIndex = find_element_index(lines, 'Size Code')
    labelIndex = find_element_index(lines, 'Label')
    ctnsIndex = find_element_index(lines, 'CTNS')
    propertyNumber = 6

    filterLength = ctnsIndex + 1
    filterLines = []
    for i in range(len(lines)):
        if i % filterLength == arrivalDateIndex or i % filterLength == varietyCodeIndex or i % filterLength == packCodeIndex or i % filterLength == sizeCodeIndex or i % filterLength == labelIndex or i % filterLength == ctnsIndex:
            filterLines.append(lines[i])

    # Calculate the number of elements we want to keep
    num_elements_to_keep = (
        len(filterLines) // propertyNumber) * propertyNumber

    # Slice the list to the desired length
    filterLines = filterLines[:num_elements_to_keep]

    # with open('content.txt', 'w') as f:
    #     f.write('\n'.join(map(str, filterLines)))

    class Cherry:
        def __init__(self, arrival_date, variety_code, pack_code, size_code, label, ctns):
            self.arrival_date = transform_date(arrival_date)
            self.variety_code = variety_code

            if pack_code[-1] == 'K':
                self.pack_code = int(pack_code[:-1])
            elif pack_code[-1] == 'L':
                self.pack_code = round(int(pack_code[:-1]) * 0.45359237, 0)

            self.size_code = size_code
            if size_code[-2:] == 'RW':
                self.size_code = int(size_code[:-2])
            if size_code[-2:] == '12':
                self.size_code = int(size_code[:-2]) + 0.5

            self.label = label
            self.ctns = ctns

        def __repr__(self):
            return f"Cherry({self.arrival_date}, {self.variety_code}, {self.pack_code}, {self.size_code}, {self.label}, {self.ctns})"

    # Grouping the filterLines into chunks of propertyNumber to form each Cherry object
    chunks = [filterLines[i:i + propertyNumber]
              for i in range(propertyNumber, len(filterLines), propertyNumber)]
    chunks = combineRows.combineRows(chunks)

    # Creating the list of Cherry objects
    cherrys = [Cherry(*chunk) for chunk in chunks]
    cherrys.sort(key=lambda cherry: (cherry.label, cherry.variety_code, cherry.pack_code, cherry.size_code))
    # print(cherrys)
    return cherrys

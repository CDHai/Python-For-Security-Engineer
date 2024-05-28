import os
import re


def find_gap(numbers):
    for i in range(1, len(numbers)):
        if numbers[i] - numbers[i-1] != 1:
            return numbers[i-1] + 1
    return None
 
 
def find_and_rename(prefix, folder):
    # Create a regex pattern to match the prefix and numbering
    pattern = re.compile(r'^{}(\d{{3}})\.txt$'.format(re.escape(prefix)))
    
    # Find all files matching the pattern
    files = [file for file in os.listdir(folder) if pattern.match(file)]
    
    # Extract numbers from file names
    numbers = sorted(int(pattern.match(file).group(1)) for file in files)
    
    #Find gap
    gap = find_gap(numbers)
    
    #While there is gap, rename files, and continuously update files and numbers arr
    while gap is not None:
        print("Gap found at number:", gap)
        count = 1
        for i, file in enumerate(files):
            old_number = int(pattern.match(file).group(1))
            new_number = gap + count
            new_name = re.sub(r'(\d{3})', str(new_number).zfill(3), file)
            while os.path.exists(os.path.join(folder, new_name)):
                new_number += 1
                new_name = re.sub(r'(\d{3})', str(new_number).zfill(3), file)
            os.rename(os.path.join(folder, file), os.path.join(folder, new_name))
            print("Renamed", file, "to", new_name)
            count += 1
        files = [file for file in os.listdir(folder) if pattern.match(file)]
        numbers = sorted(int(pattern.match(file).group(1)) for file in files)
        gap = find_gap(numbers)
        
    #If there is no gap or no more gap, print    
    print("No Gap. Done!")


folder_path = "C:/Users/admin/Desktop/automate/FillingInTheGaps"
prefix = "spam"
find_and_rename(prefix, folder_path)

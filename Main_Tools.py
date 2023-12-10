from PIL import Image
import os
import glob
import xlwings as xw
import shutil

# Get the list of files in the directory
path = glob.glob('./MANGO 16 NOV  Base - Copy/*')
# Initialize the counter for renamed or removed files
file_count = 0
file_list = []
file_condition_list = []
class Worktools:
    def __init__(self, path_file, file_count):
        """Initialize the class with the file path and the file count."""
        self.files = path_file
        self.file_count = file_count
        
    def sumfiles(self):
        print(f"{  len(self.files)}\tfiles identified")

    def file_condition(self,condition_name):
        count = 0
        #counting file that match the condition
        print(f"{len(self.files)} files identified")
        for file in self.files:
            # if file.count(condition_name) > 1 in file:
            if condition_name in file:
                file_condition_list.append(file)
                # Increment the file count
                count += 1
        self.file_count = count
        # print(len(file_condition_list))
        print(f"  {count}\tfiles that match the condition")

    def convert_image_format(self, file_from, file_to):
        count = 0
        print("\n=============  >>> Convert Image")
        # _worktools.file_condition(file_from)
        #covert all image in the directory
        print(f"  {self.file_count}\tImage will be convert from .{file_from} to .{file_to}")
        while True:
            # Ask the user for confirmation before removing the files
            confirmation = input(f"Confirm to Convert files to .{file_to}? ").lower()
            if confirmation in ['yes','ya','y']:   
                for item in self.files:
                    if item in file_condition_list:
                        file_image = Image.open(item).convert('RGB')
                        file_image.save(item.replace(file_from, file_to))
                        count += 1
                    else:
                       continue
                    # self.file_count +=1
                print(f"  {count}\tFiles image Converted to .{file_to}")
                _worktools.remove_files(file_from)
                # _worktools.sumfiles()
                break
            else:
                print("\tNo file Converted")
                break
    def rename_file(self, replace_from, replace_with):
        """Rename all files that contain spaces in their names."""
        count = 0
        print(f"  {len(self.files)}\tfiles identified")
        print(f"  {len(self.files)}\tfiles will be renamed!")
        for file in self.files:
            if " " in file:
                # Open the file as an image
                file_image = Image.open(file)
                # Save the file with a new name that replaces
                file_image.save(file.replace(replace_from,replace_with))
                # Increment the file count
                count += 1
        print(f"  {count}\tfiles renamed!")

    def remove_files(self, file_condition):
        """Remove all files that contain condition in their names."""
        count = 0
        # print(f"{len(self.files)} files identified")
        # #counting file that match the condition
        # for item in self.files:
        #     if file_condition in item:
        #         # Increment the file count
        #         self.file_count += 1
        # print(f"{self.file_count} files that match the condition will be removed!")
        print("\n=============  >>> Remove")
        # _worktools.file_condition(file_condition)
        while True:
            # Ask the user for confirmation before removing the files
            confirmation = input("Are you sure to remove these files?  ").lower()
            if confirmation in ['yes','ya','y']:  
                for item in self.files:
                    if file_condition in item:
                        # Remove the file from the directory
                        os.remove(item)
                        count += 1
                print(f"  {count}\tfiles removed")
                break
            else:
                print("\tNo file removed")
                break
    
    def check_duplicates(self):
        #rename file that countain space to non-space and insert to list
        for item in self.files:
            isi_list = item.replace(" ","")
            file_list.append(isi_list)
        #Check duplicate file from file list
        for item in file_list:
            if file_list.count(item) > 1:
                #increment the file count that duplicate
                self.file_count +=1
                print(item)
        if self.file_count > 0:
            print(f"{self.file_count} File duplicates")
        else:
            print("No Duplicate file")

    def Copy_Image_BaseOn_Sku(self, path_dest, excell_file, name_sheet, range_col):
       
        path_to = path_dest
        file_excell = excell_file
        sheet_name = name_sheet
        range_col = range_col
        count = 0
        
        workbook = xw.Book(file_excell)
        Wsheet = workbook.sheets(sheet_name)
        sku_list = Wsheet.range(range_col).value
        
        print(f"{len(self.files)}\tfiles in folder")
        print(f"{len(sku_list)}\tSKU on sheet")
        while True:
            # Ask the user for confirmation before copy the files
            confirmation = input(f"Confirm copy to {path_to}? ").lower()
            if confirmation in ['yes','ya','y']:  
                for item in sku_list:
                    for file in self.files:
                        if item in file:
                            shutil.copy(file, path_to)
                            count += 1
                        else:
                            continue
                break
            else:
                print("\tNo file Copied")
                break
                
        print(f"  {count}\tfile Copied")
        print(f"  {count/4}\tSKU Copied")
            


# Create an instance of the Worktools class
_worktools = Worktools(path, file_count)
# Call the remove_files method
# worktools.remove_files(" ")
# worktools.check_duplicates()

 

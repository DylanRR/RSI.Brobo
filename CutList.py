import sys
import os
import xlwt
import colorama
from time import sleep
# from colorama import Fore, Back, Style
colorama.init()
# import curses
# import datetime
version = 1.0


def cls():
    os.system('cls')


class Stick(object):
    """Cut list for one sick of steel, list can be run multiple times"""

    def __init__(self, material_len, qty):
        self.material_len = material_len
        self.qty = qty
        self.part_list = []
        self.usable_len = material_len - 7

    def add_piece(self, part_len, part_qty):
        """"Function to add a quantity of pieces to cut list and update usable length"""
        for x in range(part_qty):
            self.part_list.append(part_len)
        self.part_list.sort()
        self.usable_len = round(self.material_len - 6 - (sum(self.part_list) + len(self.part_list) * .098), 3)


class Sheet(object):
    """"Class for building a new spreadsheet from the cut lists supplied from the user.
    Accepted arguments for instantiation are
        pno:Program Number(int)
        stock:Length of stock in inches to be used for the cut list (float)

    line_count is an attribute used to track the line number currently being worked on in the spreadsheet to determine
    the correct row to write to

    last_piece is an attribute to hold the length of the last piece of material to allow calculation of absolute
    position needed to pull the piece out of the clamp

    wb and ws instantiate a new workbook and worksheet to write the Excel data to"""

    def __init__(self, pno, stock):
        self.pno = pno
        self.stock = stock
        self.line_count = 1
        self.last_piece = 0
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet('Tabelle1')

    def new_line(self):
        """This takes care of the repetitive data that must be written to each row and it increments the line counter

         Column 0 is the axis which is always 0 for a single axis device
         Column 1 is the program line number must be sequential
         Column 5 is the output tool which is always tool 1"""
        self.ws.write(self.line_count + 2, 0, 0)
        self.ws.write(self.line_count + 2, 1, self.line_count)
        self.ws.write(self.line_count + 2, 5, 1)
        self.line_count = self.line_count + 1

    def int_sheet(self):
        """Initialize The Excel SpreadSheet with Default Fields and passes in the user defined program number"""

        self.ws.write(0, 0, 'axis')
        self.ws.write(0, 1, 'number')
        self.ws.write(0, 2, 'demand')
        self.ws.write(0, 3, 'quantity')
        self.ws.write(0, 4, 'mode')
        self.ws.write(0, 5, 'output')
        self.ws.write(1, 0, ';')
        self.ws.write(1, 1, 'This file was generated by CutList Python Script by Lewis Miller')
        self.ws.write(2, 0, 'Pno')
        self.ws.write(2, 1, self.pno)

    def drop_first(self, work_list):
        """function to process cut list with at least one piece longer than 13 inches, cuts drop first"""
        self.new_line()
        self.ws.write(self.line_count + 1, 2, 1000 * (sum(work_list) + (len(work_list) - 1) * .098))
        self.ws.write(self.line_count + 1, 3, 1)
        self.ws.write(self.line_count + 1, 4, 0)

        self.last_piece = work_list.pop()
        while len(work_list) > 0:
            self.new_line()
            self.ws.write(self.line_count + 1, 2, 0 - min(work_list) * 1000)
            self.ws.write(self.line_count + 1, 3, work_list.count(min(work_list)))
            self.ws.write(self.line_count + 1, 4, 1)

            x = min(work_list)
            while x in work_list:
                work_list.remove(x)

    def drop_last(self, work_list):
        """function to process cut list with no pieces longer than 13 inches, cuts drop last"""
        self.new_line()
        self.ws.write(self.line_count + 1, 2, self.stock - 6)
        self.ws.write(self.line_count + 1, 3, 1)
        self.ws.write(self.line_count + 1, 4, 0)
        self.last_piece = self.stock - 4 - (sum(work_list) + (len(work_list) - 1) * .098)
        while len(work_list) > 0:
            self.new_line()
            self.ws.write(self.line_count + 1, 2, 0 - min(work_list) * 1000)
            self.ws.write(self.line_count + 1, 3, work_list.count(min(work_list)))
            self.ws.write(self.line_count + 1, 4, 1)

            x = min(work_list)
            while x in work_list:
                work_list.remove(x)

    def save_sheet(self):
        """Function to save to disk, validates filename and checks for duplicate files then
        confirms if you want to overwrite"""
        save_success = False
        while not save_success:
            try:
                file_name = input('Save As File Name:').lower().strip()
                if "\\" in file_name:
                    print('\n\nFile names cannot contain a "\\"\n')
                elif file_name + ".xls" in os.listdir('.'):
                    print(f"File name {file_name}.xls already exists.")
                    conf = False
                    while not conf:
                        o_write = input("Overwrite? (y/n)\n").lower().strip()
                        if o_write == "y":
                            self.wb.save(file_name + '.xls')
                            save_success = True
                            conf = True
                        elif o_write == "n":
                            conf = True
                        else:
                            continue

                else:
                    self.wb.save(file_name + '.xls')
                    save_success = True
            except OSError:
                print('\n\nInvalid file name.\n')
        print("\nSave successful")
        sleep(2)

    def stick_change(self):
        """Adds code to retract piece from saw, setup for trim cut, and reset pusher to home position"""
        self.new_line()
        self.ws.write(self.line_count + 1, 2, 10000 * (self.last_piece + 6))
        self.ws.write(self.line_count + 1, 3, 0)
        self.ws.write(self.line_count + 1, 4, 0)
        self.new_line()
        self.ws.write(self.line_count + 1, 2, -6000)
        self.ws.write(self.line_count + 1, 3, 1)
        self.ws.write(self.line_count + 1, 4, 1)
        self.new_line()
        self.ws.write(self.line_count + 1, 2, 1000 * (self.stock - 6))
        self.ws.write(self.line_count + 1, 3, 0)
        self.ws.write(self.line_count + 1, 4, 0)

    def write_eof(self):
        """Writes the End of File line to last line in the sheet"""
        self.new_line()
        self.ws.write(self.line_count + 1, 2, 0)
        self.ws.write(self.line_count + 1, 3, 0)
        self.ws.write(self.line_count + 1, 4, 0)


def count_sticks(work_list):
    total_sticks = 0
    for x in range(len(work_list)):
        active_list = work_list[x]
        total_sticks = total_sticks + active_list.qty
    return total_sticks


def build_xls(build_list, prog_no, stock_len, job):
    """Function to build XLS file from user entered cut lists"""

    new_sheet = Sheet(prog_no, stock_len)
    new_sheet.int_sheet()

    total_sticks = count_sticks(build_list)

    for x in range(len(build_list)):
        active_list = build_list[x]
        for c in range(active_list.qty):
            if max(active_list.part_list) > 13:
                copy_list = active_list.part_list[:]
                new_sheet.drop_first(copy_list)

            elif max(active_list.part_list) <= 13:
                copy_list = active_list.part_list[:]
                new_sheet.drop_last(copy_list)

            total_sticks = total_sticks - 1

            if total_sticks > 0:
                new_sheet.stick_change()
    new_sheet.write_eof()
    
    dir_check = os.listdir('S:/BROBO')
    if str(job) not in dir_check:
        os.mkdir('S:/BROBO/' + str(job))
    os.chdir('S:/BROBO/' + str(job))
    new_sheet.save_sheet()


def get_stock():
    """determines what length stock will be used for cut job"""
    cls()
    print("\n\n\n")
    print('Stock length must be longer than 48" (Seriously why would you use this tool for something that small?)')
    print('Stock length cannot exceed 306". Lengths will be rounded to whole inches.')
    print('Usable length is calculated automatically.\n')

    while True:
        try:
            stock_length = input("What is the stock length in inches? ")
            stock_length = round(float(stock_length))

            if stock_length > 306:
                print('\nMax length is 306", please use a shorter length')
            elif stock_length < 49:
                print('\nMinimum length is 36", do you really need to use this tool?')
            elif stock_length <= 306 >= 49:
                break
        except ValueError:
            print('\nYou must enter a number only. Do not enter " for inches. Inches are assumed\n')

    return stock_length


def get_piece(usable_len):
    """Function to get and validate part length, confirms it will fit within the remaining length"""
    # TODO add validation that confirms last piece is at least 13" if no existing pieces are at least 13" long

    while True:
        try:
            part_len = round(float(input("\nPiece length in decimal inches\n")), 3)
            if part_len == 0:
                return 0
            elif part_len >= usable_len:
                print('\nPart is too long to fit, please use a shorter length or end the list.')
                continue
            else:
                return part_len
        except ValueError:
            print('\nYou must enter a number, do not use " in the number, inches are assumed')


def get_qty(usable_len, part_len):
    """Function to get and validate quantity of parts, confirms they will fit within the remaining length"""
    while True:
        try:
            part_qty = int(input('Quantity\n'))
            if part_qty == 0:
                print('\nYou must enter a quantity of at least 1')
                continue
            elif (part_qty * part_len) + (part_qty - 1) * .098 >= usable_len:
                print('\nToo many parts, not enough remaining length, please reduce quantity')
                continue
            else:
                return part_qty
        except ValueError:
            print('\nYou must enter a number')


def format_list(parts_list):
    f = "Cut List"
    for i in parts_list:
        f = f + '\033[0;30;46m[\033[0;30;46m ' + str(i) + '"\033[0;30;46m]\033[1;37;40m'
    return f


def build_cutlist(stick_len):
    """builds list for a single cut list while verifying parts will fit within stock"""
    # TODO: Show cut list data entered so far
    cut_list = Stick(stick_len, 1)

    while True:
        cls()
        print('\033[1;32;40m\n\n\nBuilding Cut List...\033[1;37;40m\n')
        formatted_stick = format_list(cut_list.part_list)
        print(formatted_stick)
        print('Remaining usable length is ' + str(cut_list.usable_len) + '"')
        print('\nTo end list enter a length of 0')

        part_len = get_piece(cut_list.usable_len)
        if part_len == 0:
            break
        print()
        part_qty = get_qty(cut_list.usable_len, part_len)
        print('\n\n')

        cut_list.add_piece(part_len, part_qty)
        cut_list.qty = 1
    # while True:
    #    try:
    #        list_qt = int(input('\nHow many times do you want to cut this layout?'))
    #        if list_qt > 10:
    #            print('\nIt would be better to make this a single cut list')
    #            print('and have the operator run it as many times as they need')
    #            continue
    #        else:
    #            cut_list.qty = list_qt
    #            break
    #    except ValueError:
    #        print('Invalid input')
    return cut_list


def set_pno():
    """sets and validates the program number from user input"""
    cls()
    print("\n\n\nProgram Number can be any whole number from 1 - 30\n")
    while True:
        try:
            pno = int(input('Set Program Number: '))

            if 31 > pno > 0:
                return pno
            else:
                print('\nvalue must be from 1-30')
                continue
        except ValueError:
            print('\nYou must enter a number')


def set_jno():
    """Set and validates the Job Number in order to save the file to the correct path"""
    cls()
    print("Job numbers must be 4 digits long")
    while True:
        try:
            jnum = int(input('Enter Job Number: '))

            if len(str(jnum)) != 4:
                print('\nJob Numbers must be 4 digits')
                continue
            else:
                return jnum
        except ValueError:
            print('Numbers only in a job number')


def main():
    stock_len = 240
    prog_num = 1
    job_num = 0
    list_of_lists = []
    # TODO: add an option to review cut lists
    # TODO add option to edit cut lists
    # TODO add ability to read in an existing spreadsheet program to be edited

    while True:
        cls()
        print('\nCut List Builder v' + str(version))
        print('\033[1;37;40m \n\n\n.-------==\033[0;34;47m Menu \033[1;37;40m==-------.')
        print('1: Set Stock Length [\033[1;32;40m ' + str(stock_len) + '"\033[1;37;40m ]')
        print('2: Set Program #    [\033[1;32;40m ' + str(prog_num) + '\033[1;37;40m ]')
        print('3: Set Job #        [\033[1;32;40m ' + str(job_num).zfill(4) + '\033[1;37;40m ]')
        print('4: Add cut list to program\n5: Compile program to XLS\n6: Start New Program\n7: Quit\n')
        menu_choice = input("What do you want to do?:")

        if menu_choice == "1":
            stock_len = get_stock()
            continue
        elif menu_choice == "2":
            prog_num = set_pno()
            continue
        elif menu_choice == "3":
            job_num = set_jno()
            continue
        elif menu_choice == "4" and stock_len == 0:
            print('\nYou must set the stock length before defining any cut lists')
            continue
        elif menu_choice == "4":
            print("\nUser defines cut list")
            cut_list = build_cutlist(stock_len)
            list_of_lists.append(cut_list)
            continue
        elif menu_choice == "5" and len(list_of_lists) == 0:
            print('\nNo cut lists have been defined yet')
            continue
        elif menu_choice == "5" and prog_num == 0:
            print('\nNo program number has been set')
            continue
        elif menu_choice == "5" and job_num == 0000:
            print('\nNo job number has been set')
            continue
        elif menu_choice == "5":
            build_xls(list_of_lists, prog_num, stock_len, job_num)
            continue
        elif menu_choice == "6":
            while True:
                verify_delete = input('\033[0;37;41mAre you sure you want to delete the lists in memory?\033[1;37;40m')
                if verify_delete.lower() == "y" or verify_delete.lower() == "yes":
                    list_of_lists = []
                    prog_num = prog_num+1
                    break
                elif verify_delete.lower() == "n" or verify_delete.lower() == "no":
                    break
                else:
                    print("Yes or No?")
                    continue
            continue

        elif menu_choice == "7":
            while True:
                verify_quit = input('\nAre you sure you want to quit? ')
                if verify_quit.lower().strip() == "y" or verify_quit.lower().strip() == "yes":
                    sys.exit(0)
                elif verify_quit.lower().strip() == "n" or verify_quit.lower().strip() == "no":
                    break
                else:
                    print("Yes or No?")
                    continue
        else:
            print("Invalid input")
            continue

    # int_sheet()
    # stock_len = get_stock()
    # cut_list = build_cutlist(stock_len)
    # print(cut_list)


if __name__ == "__main__":
    main()

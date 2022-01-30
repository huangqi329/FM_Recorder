# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import sys
import FileParser.JsonParser as jp

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    filePath = sys.path[0]
    if len(sys.argv) > 1:
        filePath = sys.argv[1]
    else:
        filePath = './/'

    jp.scanAllFiles(filePath)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

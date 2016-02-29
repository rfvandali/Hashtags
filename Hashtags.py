######################################### LIBRARIES TO IMPORT ########################################################
import re                           # Regular expressions library
from collections import Counter     # Import counter function
import xlwt                         # Library to write to excel spreadsheet



######################################### DEFINITION OF FUNCTIONS #####################################################
#### Function used to find a whole word in a string. Eg. findWholeWord(wordtobefound)(stringtosearch)
def findWholeWord(w):
    return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

#### Function used to count number of letters in string
def count_letters(word):
    return len(word) - word.count(' ')



################################################## MAIN CODE ##########################################################
#### Stores words from doc1 to doc6 in mega string, tallies all the words and then stores the 10 most common
cap_words = []
for i in range(1,7):                                            # Searches through all six text documents
    doc = 'doc' + str(i) + '.txt'                               # Updates current txt document
    with open(doc, encoding="utf8") as doc:
        passage = doc.read()
    words = re.findall(r'\w+', passage, flags=re.IGNORECASE)    # Finds words from current document
    cap_words = [word.upper() for word in words] + cap_words    # Updates mega string
word_counts = Counter(cap_words).most_common(10)                # Arbitrarily chosen 10 most common (can be changed)

#### Makes the excel table look pretty
book = xlwt.Workbook(encoding="utf-8")                          # Creates a new workbook
sheet1 = book.add_sheet("Sheet 1")                              # Create a new sheet (called Sheet 1)
sheet1.write(0, 0, "Word")                                      # Name of first column
sheet1.write(0, 1, "Number of occurrences")                     # Name of second column
sheet1.write(0, 2, "Documents")                                 # Name of third column
sheet1.write(0, 3, "Sentences containing the word")             # Name of fourth column
style = xlwt.XFStyle()
style.alignment.wrap = 1                                        # Defines 'style' for wrapping text in cells
fourth_col = sheet1.col(3)                                      # Defines 4th column
fourth_col.width = 256 * 60                                     # 4th column 60 characters wide (-ish)
third_col = sheet1.col(2)                                       # Defines 3rd column
third_col.width = 256 * 15                                      # 3rd column 15 characters wide (-ish)
second_col = sheet1.col(1)                                      # Defines 2nd column
second_col.width = 256 * 25                                     # 2nd column 25 characters wide (-ish)

#### Main part of executing code
for k in range(len(word_counts)):                               # Pass through each of the top 10 most common words

    new_line = 0                                                # Dummy counter used to add two returns between sentences
    sentence_containing_word = ''                               # Initiliase sentence_containing_word to be empty
    max_char_achieved = 0                                       # Dummy counter to know when the maximum characters for a cell are achieved
    excel_word = re.sub(r'\d+', '', str(word_counts[k]))        # Removes numbers from most common word list
    excel_word = re.sub(r'\W+', '', excel_word)                 # Removes non unicode characters for most common word list (only want alphabet)
    number_occur = re.sub(r'\D+', '', str(word_counts[k]))      # Removes characer from most common word list (only want numbers)
    sheet1.write(k+1, 0, excel_word)                            # Prints most common word to spreadhseet
    sheet1.write(k+1, 1, number_occur)                          # Prints number of occurrences of most common word to spreadsheet
    comma = 0                                                   # Dummy counter used to make table pretty (adds comma if more than one document listed)
    found_doc = ''                                              # initialise

    for i in range(1,7):                                        # Pass through each file (doc1.txt to doc6.txt)
        doc = 'doc' + str(i) + '.txt'                           # Updates current txt file
        text_file = open(doc,encoding="utf8")
        sentence = text_file.read().split('.')                  # Assigns each sentence in txt file to list/array

        for j in range(len(sentence)):                          # Pass through each sentence in txt file
            if max_char_achieved == 0:
                sentence[j] = sentence[j].strip(" ")                        # Makes the text pretty (removes whitspace before sentence)
                sentence[j] = sentence[j].strip("\n")                       # Makes the text pretty (removes new lines at end of sentence)
                sentence[j] = sentence[j] + '.'                             # Makes the text pretty (removes new lines at end of sentence)

                if findWholeWord(excel_word)(sentence[j]):                  # Finds the sentence containing common word by checking each sentence (if too many characters, stop search)
                    if count_letters(sentence_containing_word) > 100:       # Arbitrarily have stopped recording sentences at 100 characters. Can be more but becomes unweildly
                        max_char_achieved = 1
                    if new_line == 1:
                        sentence_containing_word = sentence[j] + '\n\n' + sentence_containing_word
                    else:
                        sentence_containing_word = sentence[j]
                        new_line = 1

        with open(doc, encoding="utf8") as doc:                 # Looks for a match of common word to identify which document(s) it came from
            for line in doc:
                if findWholeWord(excel_word)(line):             # Looks for a match of common word to identify which document(s) it came from
                    if comma == 1:                              # if statement used to essentially make the table look pretty
                        found_doc = 'doc'+str(i)+'.txt\n' + found_doc
                        break                                   # break because once a single match is found in the document, time to move on
                    else:
                        found_doc = 'doc'+str(i)+'.txt'
                        comma = 1
                        break                                   # break because once a single match is found in the document, time to move on

    sheet1.write(k+1, 2, found_doc, style)                      # Writes documents where the word is found to spreadsheet

    sheet1.write(k+1, 3, sentence_containing_word,style)        # Writes the sentences where the word is found to spreadsheet


book.save("Hashtags.xls")                                                      # Save workboook with name 'hashtags




def whiteSpaceAdder(text):
    if(len(text) >= 5):
        # do nothing
        print("More than 24")
    else:
        # add empty line
        print("Less than 24")

print('Enter some text: ')
text = input()
print('Checking..')
print('')
whiteSpaceAdder(text)
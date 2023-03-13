import pandas as pd

# Read the Excel file into a pandas dataframe
df = pd.read_excel('Input.xlsx', sheet_name='Sheet1')

# Filter out rows containing the string "response"
df = df[~(df['Payload Type'].str.contains(r'Response', na=True))]


# convert the dataframe to a list of dictionaries
data = df.to_dict('records')

# remove first row
data = data[1:]

# function to wrtie this data to a new text file with added custon text
def write_to_file(data):
    
    #loop through the data and write to file
    with open("Output.txt", "w") as file:
        for row in data:
            file.write("Name of the service: " + row['Payload Type'] + " \n")


# call the function
write_to_file(data)

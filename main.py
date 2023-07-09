import openpyxl
import openai

openai.api_key = 'API_KEY'

SP500Book = openpyxl.load_workbook('SP500YTD.xlsx')
SP500YTDSheet = SP500Book.active
# SP500Book['specificSheetName']

data = []
for row in SP500YTDSheet.iter_rows(values_only=True):
    data.append(row)

formattedData = ""

for row in data[3:]:
    sentence = f'{row[0]}: STOCK RETURN = {row[1]}%, TOTAL RETURN = {row[2]} \n'
    formattedData += sentence

# the beginning of every (user) prompt
baseMessage = "Provide insights on the trends seen in the following economic data from 2020: \n"

# telling GPT what character it is to play
systemMessage = "You are a senior S&P 500 analyst that has been following global and political trends and their \
influence on the S&P 500 for 25 years. Your manner is polite but extremeley concise. You always answer with \
specific facts about the state of the world and how different markets are moving."

print("Prompt sent to GPT: \n" + baseMessage + formattedData)

completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "system", "content": systemMessage},
        {"role": "user", "content": baseMessage + formattedData}
    ]
)

print(completion.choices[0].message['content'])


# Code for continuous conversation

def getGPTResponse(messageStream):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messageStream
    )
    return response.choices[0].message


running = True
system_msg = {"role": "system", "content": systemMessage}
data_msg = {"role": "user", "content": baseMessage + formattedData}
msgStream = [system_msg, data_msg]

while running:
    i = input("(R) reply (Q) quit \n")
    if i == "R":
        reply = input("Please ask any clarifying or further questions.")
        print("Processing Response...")
        user_msg = {"role": "user", "content": reply}
        msgStream.append(user_msg)
        assistant_response = getGPTResponse(msgStream)
        msgStream.append(assistant_response)
        print(assistant_response['content'])
    elif i == "Q":
        running = False


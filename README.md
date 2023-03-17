# Add GPT Chatbot to Microsoft Word
Create a powerful chatbot in Microsoft Word using ChatGPT
<br /><br />

<!-- badges: start -->
[![Awesome](https://awesome.re/badge.svg)](https://awesome.re)&nbsp;&nbsp;
![Lifecycle:Stable](https://img.shields.io/badge/Lifecycle-Stable-97ca00)&nbsp;&nbsp;
![](https://img.shields.io/badge/Maintained%3F-yes-green.svg)&nbsp;&nbsp;
[![CI](https://github.com/analyticsinmotion/add-gpt-chatbot-to-microsoft-word/actions/workflows/blank.yml/badge.svg)](https://github.com/analyticsinmotion/add-gpt-chatbot-to-microsoft-word/actions/workflows/blank.yml)&nbsp;&nbsp;
[![MIT license](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/analyticsinmotion/add-gpt-chatbot-to-microsoft-word/blob/main/LICENSE.md)&nbsp;&nbsp;
![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)&nbsp;&nbsp;
![Microsoft Word](https://img.shields.io/badge/Microsoft_Word-2B579A?logo=microsoft-word&logoColor=white)&nbsp;&nbsp;
<!-- badges: end -->

<!-- DESCRIPTION -->
## 1. Description

Easily access ChatGPT's awesome chatbot capabilies in Microsoft Word. This application captures the conversation history between a user and the chatbot. By utilizing conversation history, the chatbot can mimic an awareness of context and thus provide responses that make more sense.
<br /><br />




https://user-images.githubusercontent.com/52817125/225483447-b59057f0-8cd8-4a75-9062-a0fb75ccf5ac.mp4




<br />

<!-- GETTING STARTED -->
## 2. Getting Started
### 2.1 Dependencies
- Requires an OpenAI API Key (create an account and get API Key at <a href="https://chat.openai.com">https://chat.openai.com</a>)
- Requires Microsoft Windows 10/11 (<a href="https://www.microsoft.com/en-au/windows">https://www.microsoft.com/en-au/windows</a>)
- Requires Microsoft Word 365 (<a href="https://www.microsoft.com/en-us">https://www.microsoft.com/en-us</a>)

Please be aware of the [costs](https://openai.com/pricing) associated with using the OpenAI API when utilizing this project.


### 2.2 AI Models

This application uses the following OpenAI model:
 
| Model  | Description | Max tokens | Training data |
| ------------- | ------------- |------------- | ------------- |
| gpt-3.5-turbo  | Most capable GPT-3.5 model and optimized <br /> for chat at 1/10th the cost of text-davinci-003. <br /> Will be updated with our latest model iteration. | 4,096 tokens | Up to Sep 2021 |

Further information about all OpenAI models can be found here: <a href="https://platform.openai.com/docs/models/overview">https://platform.openai.com/docs/models/overview</a>

We endeavour to test and integrate newer models when they are become Generally Available (GA). Models released as a 'Limited Beta' will not be integrated until they become GA.


<br />
    
<!-- INSTRUCTIONS -->
## 3. Instructions

  - To start a chat, write a message anywhere in Microsoft Word
  - Select your message and click the *Chat Completion* button in the AI Assistant tab
  - Wait for the model to respond
  - The chat response will appear under initial message
  - Repeat the steps above to continue the chat 
  <br />
  
  To chat about a new topic simply click the *Chat Reset* button in the AI Assistant tab 
  
<br />

### 3.1 Chat Completion Example 1

User Message #1
```
Who won the world series in 2020?
```

Chatbot Reply
```
The Los Angeles Dodgers won the World Series in 2020.
```

User Message #2
```
Where was it played?
```

Chatbot Reply
```
The 2020 World Series was played in Arlington, Texas at Globe Life Field, the home stadium of the Texas Rangers.
```
<br />

<strong>Conversation Flow</strong>
> User: Who won the world series in 2020?
> > Chatbot: The Los Angeles Dodgers won the World Series in 2020.

> User: Where was it played?
> > Chatbot: The 2020 World Series was played in Arlington, Texas at Globe Life Field, the home stadium of the Texas Rangers.



<br />

<!-- Installation -->
## 4. Installation

There are 4 basic steps in order to add a ChatGPT button into Microsoft Word:
  1. Enable the Developer Tab
  2. Import the VBA script files
  3. Create the Chat Completion and Chat Reset buttons 
  4. Add your OpenAI APIKey

Each of these steps are fully outlined below. 
<br /><br />

### 4.1 Enable the Developer Tab

The Developer tab isn't displayed by default, but you can add it to the ribbon.

**Step 1** - On the File tab, go to Options > Customize Ribbon.

**Step 2** - Under Customize the Ribbon and under Main Tabs, select the Developer check box.

<img src=".github/assets/images/enable-developer-tab-highlighted.png" width=100% height=100%>
<br />

The latest instructions to enable the Developer Tab from Microsoft can be found here: 
<a href="https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792">https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792</a>
<br /><br />

### 4.2 Import the Chat.bas and ChatReset.bas files

**Step 1** - Download and Save the latest Chat.bas and ChatReset.bas file from the src/windows folder in this repository.
<br />

Keep the location of where the file is saved as you will need it later.<br />
<br />

**Step 2** - On the Developer tab, click the Visual Basic button.

<img src=".github/assets/images/developer-tab-visual-basic.png" width=100% height=100%>
<br />

**Step 3** - On the File tab, go to Import File...

<img src=".github/assets/images/visual-basic-file-import-section.png" width=100% height=100%>
<br />

**Step 4** - Add the two .bas files
 - Select the Chat.bas file and click Open
 - Select the ChatReset.bas file and click Open
<br /><br />

### 4.3 Add your Chat Completion and Chat Reset buttons into the Microsoft Word Ribbon

*Please Note:* This project closely relates to two of our other projects: 
 - **Add ChatGPT to Microsoft Word** - project found <a href="https://github.com/analyticsinmotion/add-chatgpt-to-microsoft-word">here</a> 
 - **Create ChatGPT Images in Microsoft Word** - project found <a href="https://github.com/analyticsinmotion/create-chatgpt-images-in-microsoft-word">here</a> 

If you have have already added one or both of these projects into Microsoft Word you can start at Step 3 of this section.
<br /><br />

**Step 1** - Add a new tab
<br />
  - On the File tab, go to Options > Customize Ribbon
  - Click New Tab
<br />

<img src=".github/assets/images/options-customize-ribbon-new-tab.png" width=40% height=40%>
<br />

 **Step 2** - Rename the New Tab to **AI Assistant**

<img src=".github/assets/images/options-customize-ribbon-rename-tab.png" width=35% height=35%>
<br />

**Step 3** - Rename New Group (Custom) to **ChatBot**

<img src=".github/assets/images/rename-new-group-chatbot.png" width=35% height=35%>
<br />

**Step 4** - Select Macros in the Choose Commands from dropdown box

<img src=".github/assets/images/choose-commands-from-macros.png" width=35% height=35%>
<br />

**Step 5** - Select the Chat Macro and click Add >>

<img src=".github/assets/images/add-the-chat-macro-into-new-group.png" width=75% height=75%>
<br />

**Step 6** - Rename button to **Chat Completion**, select a Symbol and click OK

<img src=".github/assets/images/rename-button-to-chat-completion.png" width=35% height=35%>
<br />

**Step 7** - Repeat Steps 5 & 6 to to create a **Chat Reset** button
 - Ensure you select the ChatReset Macro to link to your Chat Reset button in this step
<br />





After the preceding steps have been completed the Microsoft Word screen should look like the following:

<img src=".github/assets/images/screen-after-chatbot-buttons-added.png" width=100% height=100%>
<br />


### 4.4 Add your OpenAI APIKey into Windows

**Step 1** - Open the Start menu and start typing "environment variables". When the best match appears click "Edit the system environment variables" result.

<img src=".github/assets/images/add-env-var-step-1.png" width=75% height=75%>
<br />

**Step 2** - Click the "Environment variables" button under the "Advanced" tab.

<img src=".github/assets/images/add-env-var-step-2.png" width=50% height=50%>
<br />

**Step 3** - Create a new user variable by clicking "New" under the "User Variables" section.

<img src=".github/assets/images/add-env-var-step-3.png" width=50% height=50%>
<br />

**Step 4** - Type the variable name **OPENAI_API_KEY** in the first field and your OpenAI APIKEY in the variable value field. Then click OK.

<img src=".github/assets/images/add-env-var-step-4.png" width=50% height=50%>
<br />

**Step 5** - **IMPORTANT** You must restart Windows to apply the new environment variable
<br /><br />


<!-- Best Practices for API Key Safety -->
## 5. Best Practices for API Key Safety

Your OpenAI APIKEY key/s should be kept secure and private at all times.

Please follow the best practices guide for API security from OpenAI 
<br />
<a href="https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety">https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety</a>

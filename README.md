# Add GPT Chatbot to Microsoft Word
Create a powerful chatbot in Microsoft Word using ChatGPT
<br /><br />

<!-- badges: start -->

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
<br /><br />
    
<!-- INSTRUCTIONS -->
## 3. Instructions

  - To start a chat, write a message anywhere in Microsoft Word
  - Select your message and click the *Chat Completion* button in the AI Assistant tab
  - Wait for the model to respond
  - The chat response will appear under initial message
  - Repeat the steps above to continue the chat 
  <br />
  
  To chat about a new topic simply click the *Chat Reset* button in the AI Assistant tab
  
  
<br /><br />

### 3.1 Text Completion Example 1


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


<br /><br />




<!-- Best Practices for API Key Safety -->
## 5. Best Practices for API Key Safety

Your OpenAI APIKEY key/s should be kept secure and private at all times.

Please follow the best practices guide for API security from OpenAI 
<br />
<a href="https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety">https://help.openai.com/en/articles/5112595-best-practices-for-api-key-safety</a>

# crunched-kiss-
a KISS version of crunched that AI chat with your excel 


Hi! This exercise is a way for us to understand how you work as a full-stack engineer. Building production-quality software is a process of making trade-offs, structuring code for maintainability, and solving problems creatively. We don't expect a polished product in four hours. We do expect you to show us how you think, how you structure code, and how resourceful you are with limited time. If you have to choose between more features and good code, - choose good code. Include any notes, diagrams, or documentation that showcase your process. If you have questions, contact your hiring point of contact.
Time limit: 4 hours. We will know if you spent significantly more.
What is Crunched?
This is Crunched inside Excel as of December 2025:
A screenshot of a computer

AI-generated content may be incorrect.
Crunched lives in Excel as an add-in downloaded from Microsoft app-source (Add-ins button). It has programmatic access to any functionality a normal Excel user would have, as well as other features, such as the ability to search the web or interact with some of the applications that users typically use in their day-to-day work. Typical Crunched use-cases are building full financial models from scratch, iterating on large existing models, researching autonomously, or finding errors and potential issues in workbooks.
For reference, the actual Crunched stack is:
Frontend: TypeScript, React, Office.js
Backend: Python, FastAPI, LangGraph
 
The task
Build a simplified version of Crunched.
At minimum, a user should be able to open a chat interface in Excel's task pane, send a message to an AI agent, and have the agent read from and write to the spreadsheet. It must be able to interface with workbooks of any size.
How you accomplish this is up to you. Use whatever stack, architecture, and approach you think is best. We care about your reasoning, not a specific implementation pattern.
 
Getting Started
To scaffold an Excel add-in with React, you can use Microsoft's quickstart:
https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-react
This is optional. Use whatever setup works for you.
Here is an Anthropic API key. Use it as much as you want: xxxx[will copy and paste later]
 
Pitfalls
Office add-ins run inside a sandboxed WebView controlled by Excel. This WebView enforces strict security: it will not load any content over plain HTTP. Every URL in the manifest, the task pane, icons, commands, must be HTTPS, and the certificate must be trusted by the operating system. This means local development needs two separate trusted certificates: one for the webpack dev server (frontend) and one for the FastAPI/uvicorn server (if you choose to use that).
If you’re building any Office add-in with a local dev server, the trick is the same: use mkcert to create certificates signed by a local CA that gets installed into your OS trust store.
brew install mkcert
mkcert -install                           # creates a local CA and trusts it in macOS Keychain
mkcert localhost 127.0.0.1 ::1         # generates cert + key files
Then point your dev server (webpack, vite, express, uvicorn, whatever) at the generated .pem files. Excel’s WebView will now trust your local HTTPS server. If you use Excel online you might be able to skip this problem, but if you can make it work quickly with desktop Excel, that’s the recommended solution.
Submission
1.    Create a GitHub repository with your solution
2.    Include a README.md with setup instructions and general thoughts
3.    Send the repository link to your hiring contact or recruiting@usecrunched.com
If the repo is private, add access for: markusskagemo and larsgmu
Be ready to walk us through your work in a 15-minute call.

Picture description:
The right panel is an AI assistant sidebar called "Crunched" — an "AI analyst in Excel" chat pane. It reads:

Header: "Crunched" / "No new conversation"

Intro message:Hi, I'm Crunched — your AI analyst in Excel. I can help you with things like:
Error checking and fixing models
Building financial or business models
Analyzing and comparing scenario models
Pulling insights from the web and analyzing them
What should we work on first?
(6:31 PM)

Suggested prompts (chips/messages below):
Error check Budget model worst case (partially legible)
Check assumptions in Revenue build-up model
Create comparison dashboard across scenarios (partially legible)


Bottom: input box "Ask me anything..." with "Reason" and "Agent Mode" toggles.

So it's basically an Excel-embedded copilot for financial model work — error-checking, scenario comparison, dashboards — sitting next to what looks like a multi-scenario budget/revenue model spreadsheet. Heads-up: the screenshot is low-res, so the three suggestion lines are my best read; everything else is verbatim.

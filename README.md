# Project Title

PresenterPal

## About The Project

The **PowerPoint Pal** is a dynamic add-in for Microsoft PowerPoint that empowers users to apply formatting changes to slides using natural language commands. Leveraging the capabilities of OpenAI's GPT-3.5-turbo model, this add-in interprets user instructions and applies precise formatting to specific shapes across multiple slides seamlessly.

## Starting Up

### Dependencies

Prerequisites
Node.js (LTS version recommended): Download Node.js
npm (comes with Node.js)
Microsoft PowerPoint (Desktop version)
Setup Steps
Clone the Repository

### Installing

bash
```
Copy code
git clone https://github.com/yourusername/your-repo-name.git
cd your-repo-name
```

bash
```
Copy code
npm install
```

Set Up Environment Variables

Create a .env file in the project root:

bash
```
Copy code
touch .env
```
Add your OpenAI API key to the .env file:

env
```
Copy code
OPENAI_API_KEY=your-openai-api-key-here
```
Note: Replace your-openai-api-key-here with your actual OpenAI API key. Do not share this key publicly or commit it to version control.

Trust the Self-Signed Certificate (For Local Development)

Install and trust the certificates:

bash
```
Copy code
npx office-addin-dev-certs install
```
Follow the prompts to install and trust the certificates required for local development.

Build and Start the Project

```
bash
Copy code
npm run build
npm start
```
This will start the local server and open PowerPoint with the add-in sideloaded.
Sideload the Add-in Manually (If Necessary)

If the add-in doesn't load automatically:

In PowerPoint, go to Insert > My Add-ins > Manage My Add-ins.
Select Upload My Add-in > Browse.
Choose the manifest.xml file from the project directory.
Use the Add-in

In PowerPoint, open the add-in's task pane.
Enter a command in the input field (e.g., Create a slide about the Roman Empire).
Click Execute Command to see the results generated by the OpenAI API.
Security Reminder
Never expose your OpenAI API key in client-side code or commit it to any repository.
Keep the .env file private and consider adding it to your .gitignore file.

### Executing program

```
npm install
npm run dev-server
```

## Authors

Contributors names and contact info

Justin Pardo
Carlos Lopez  
Jose Jimenez
Anthony Montelongo - Navejar

## Version History

* 0.1
    * Initial Release

## License

This project is licensed under the Source-Available Licenses - see the LICENSE.md file for details

## Acknowledgments

Inspiration, code snippets, etc.
* [API reference documentation]((https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office))
* [OpenAI API](https://platform.openai.com/docs/api-reference/introduction)

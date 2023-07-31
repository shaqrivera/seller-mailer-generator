![license badge](https://img.shields.io/badge/license-Apache_License_2.0-blue)
  # Seller Email Generator
  ## Description
  This project takes a dynamic list of property owners from a .csv file and creates marketing letters for each owner.
  ## Table of Contents
  - [Installation](#installation)
  - [Usage](#usage)
  - [Contributing](#contributing)
  - [Tests](#tests)
  - [Questions](#questions)
  - [License](#license)
  ## Installation
  After downloading the project to your local system, use the terminal to navigate to the directory where the project is saved. Execute the following command `npm i` in order to install all the required dependencies for the project.

  
  <img width="689" alt="Screenshot 2023-07-29 at 4 43 20 PM" src="https://github.com/shaqrivera/seller-mailer-generator/assets/90933707/f8f8e96a-b17d-426c-ac34-5ff67c215acf">

  ## Creating Templates
Creating a template is simple. Any `.docx` file can be used as a template. Simply add `{#loop}` to the beginning of your template, and `{/loop}` at the end of your template. Anywhere within the template that you want to use data from the `.csv` file, simply include the name of the data column you'd like to display between brackets in your template. So for example, to include the owner's full name in your template, add `{owner}` to the template where you would like it to appear.

### Full list of supported data columns
 - Owner -> `{owner}`
 - Mailing Address 1 -> `{mailingAddress1}`
 - Mailing Address 3 -> `{mailingAddress3}`
  

  ## Usage
  Once all the dependencies are installed, the project is ready to use. Inside the project directory, save your desired .csv file with the name `seller-data.csv`. Save your template for the letters with the name `letter-template.docx`. Save your template for the envelopes with the name `envelope-template.docx`.


<img width="917" alt="Screenshot 2023-07-30 at 6 23 44 PM" src="https://github.com/shaqrivera/seller-mailer-generator/assets/90933707/5462b1bc-4266-4036-a75c-7883b0c76df5">
  
  
  After ensuring you've saved your files with the correct names in the correct directory, use the terminal to execute the following command `npm start`. 

  
  <img width="689" alt="Screenshot 2023-07-29 at 4 44 08 PM" src="https://github.com/shaqrivera/seller-mailer-generator/assets/90933707/488832f5-6194-43a1-96d7-4349eac596ca">
  

  If everything is successful, you should see a message in the terminal, and there should be a newly created file, named `generated-seller-letters.docx` in the project directory.

  <img width="917" alt="Screenshot 2023-07-30 at 6 39 50 PM" src="https://github.com/shaqrivera/seller-mailer-generator/assets/90933707/5660d8bb-ce9a-4ccd-8f97-c398f924f8c4">

  
  ## Contributing
  If you have any feature requests for this, just contact me.
  ## Tests
  N/A
  
  ## Questions
  Github username : <a href="https://github.com/shaqrivera">shaqrivera</a>
  
  If you have any questions, please submit inquiries to <a href="mailto:shaq.rivera@gmail.com">shaq.rivera@gmail.com</a>.
  
  ## License
  This project is using the Apache License 2.0. For more information, refer to following link.
    [Apache License 2.0](https://spdx.org/licenses/Apache-2.0.html)
  ---

# Project report helper MacOS

This programme helps a user collect data and documents required to generate project reprots for an ATHENA Lesson

## How to use

### Step 0 Prepare your environment
 - All Python packages are listed in seta-button-environment.yml
 - It helps to have Anaconda so you run `conda env create -f seta-button-environment.yml` to install all packages.
 - most conversions to pdf require *wkhtmltopdf 0.12.5 (with patched qt)* to effectively convert Notebooks to pdf

### Step 1 prepare external files
- Download all zip files submitted to athena and place them inside a directory called `submissions/`
- [Setup the predict sheet](https://www.notion.so/explore-ai/How-do-I-create-Project-documents-for-group-work-f44f736d8b7b4fbca40be8c09679665c)
- [Get a Google API credentials file](https://developers.google.com/workspace/guides/create-credentials) and connect it to your Google Sheet[A medium article, helpful at time of writing](https://medium.com/@a.marenkov/how-to-get-credentials-for-google-sheets-456b7e88c430)

### Step 2 setup
1. Install / clone into the directory where you intend to generate the project reports
2. open a terminal
3. run `./setup.sh` to generate the directory structure, and copy the submissions directory if it exists. 

- [optional] NBconvert's templates make calls to the Mathjax Github repo, and this programme quickly uses up calls-per-minute quotas if you have to re-run it. [Install and configure the NBconvert custom template](https://nbconvert.readthedocs.io/en/latest/customizing.html) (It's a bit confusing. [here are the official documents.](https://nbconvert.readthedocs.io/en/latest/customizing.html). Because this requires edits to the NBconvert config file, [Be sure to also read the chapters about generating the config file.](https://nbconvert.readthedocs.io/en/latest/config_options.html) If you get stuck, [Here is a practical example](https://stackoverflow.com/questions/64127278/what-is-the-proper-way-to-specify-a-custom-template-path-for-jupyter-nbconvert-v) that can help you configure the nbconvert configuration files). 
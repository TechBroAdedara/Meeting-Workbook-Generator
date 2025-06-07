# Meeting-Workbook-Generator
A midweek meeting schedule generator for Jehovah's Witnesses. Niche project I decided to build because my father always bothered me to prepare a schedule for him. 

It grabs the content of the [Watchtower Online Library website](https://wol.jw.org/en/wol/library/r1/lp-e/all-publications/meeting-workbooks/life-and-ministry-meeting-workbook-2025), checks the month you specify, and loops through each of the weeks, parsing it and creating a table in a Word document for it.

The usual programmer meme—spending 2 hours automating what you could have done manually in 10 minutes—applies here. Except this was built in 48 hours, nonstop. 
I would usually spend about 1.5 hours getting this done manually, but this script gets it done for me in 20-ish seconds. Yup, 20 seconds. A 99.63% decrease in the time I would have spent without this script. 

I was too determined to get this done. Now that it is finished, I can save multiple hours of my time.

For some of you reading this, this script might not be too useful for you, but its purpose for being on GitHub is just to showcase how webscraping, as well as ``python-docx`` works.

# Tools used
- ``HTTPX``: For requests
- ``BeautifulSoup``: For parsing the web pages
- ``Python-docx``: Library for generating the docx (Microsoft Word) files.
- ``Icecream``: Used to pretty-pretty print data for debugging. (You should try it. It's lovely)

# How to get started
- Clone this repo into your local machine using
  ```bash
  git clone [URL]
  ```
- Install all requirements for the package using
  ```bash
  pip install -r requirements.txt
  ```
- It is a terminal app. So navigate to the ``Workbook_Gen`` file and click on ``Run``
- Enter the month of the schedule you want to generate and watch it do its magic.
- A new file will be generated in ``GeneratedDocs/`` for you.
 
# Thank you!

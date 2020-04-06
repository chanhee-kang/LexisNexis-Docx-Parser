## LexisNexis Docx parser
LexisNeixs Docx parser is a tool that parses word(docx) format of LexisNexis's World Major Publication

### Set up
Download this git for initiating this program.
```
$git clone https://github.com/chanhee-kang/LexisNexis-Docx-Parser.git
```
Then, you have to download docx file(World Major Publication) from LexisNexis.<br><br>
The docx file looks like as following picture:
<img width="654" alt="스크린샷 2020-04-04 오후 5 54 21" src="https://user-images.githubusercontent.com/26376653/78422856-62567e00-769d-11ea-86af-e1a0d408b9de.png">

Also, 'docx2txt' doesn't contain in Anaconda so you need to...
```
$pip install docx2txt
```
If you don't use Anaconda for python, then you also need to install as following:
```
$pip install docx2txt
$pip instal pandas
$pip instal os
```
### Start
Put your docx file directory
```
text = docx2txt.process("*.docx")
```
Following line is for matching countries for the articles depends on the csv file.
```
country = pd.read_csv('mwp_list.csv')
```
### Limitation
1. Country will be shown as unknown if there is no matching country in mwp_list
2. Loaded as a single file only
### Contact
If you have any requests, please contact: [https://ck992.github.io/](https://ck992.github.io/).

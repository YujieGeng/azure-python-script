<small><i><a href='http://ecotrust-canada.github.io/markdown-toc/'>Table of contents generated with markdown-toc</a></i></small>

# Usage:

## [getCaseIDList.py](https://github.com/YujieGeng/azure-python-script/blob/main/getCaseIDList.py)
## Usage Examples:
### Your folder is in the same level as Inbox:
```shell
python getCaseIDList.py --folderPath "CB" 
```
### Your folder is under Inbox but has subfolders:
```shell
python getCaseIDList.py --folderPath "Inbox\\test1\\subtest" --subject "Cx Story" --outputFileName "mytestcaseid"
```
### Your folder is just under Inbox:
```shell
python getCaseIDList.py 
```



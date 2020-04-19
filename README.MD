# Excel Files Merger

## Introduction 
Merge Multiple Excel Files into a Single File

## Prerequisites
- Python 3

## Installation
1. Clone the repository
    ```
    git clone https://github.com/harshanas/excel-files-merger.git
    ```
2. Install dependencies
   ```
   pip install -r requirements.txt
   ```

## Usage
```
python merge.py d f s
```
| Parameter        |Description | Example |
|: -------------: |-------------| -------------| 
| d      |  Files Directory to be merged | ```H:\Work\Scraped```
| f     |  Destination File Name |  ```ALL.xlsx```
| s | Starting Row | ```1```

### Example

```
python merge.py H:\Work\Scraped ALL.xlsx 1
```


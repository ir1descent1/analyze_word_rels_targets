# Analyze Targets of Word Document .rels Files

## Motivation

Simply put, ***Follina***. Also known as [CVE-2022-30190](https://nvd.nist.gov/vuln/detail/CVE-2022-30190). This exploit demonstrated (amongst other things) that, unbeknownst to the user, a Microsoft Office document could access an arbitrary website. From there, it could pull files from the website to perform RCE on the user's system. Now, in the case of *Follina*, RCE was achieved through an exploitation of MSDT, which has since been patched. Nonetheless, it is not unreasonable for RCE to be achieved by exploitation of *some other software* on a system, especially when considering opening Office documents on a non-Windows system like Linux. 

A helpful demonstration of this exploit can be found in [John Hammonds POC video](https://www.youtube.com/watch?v=dGCOhORNKRk).

## Word Document analysis

Word documents of the ".docx" file extension are really just compressed ZIP files. Typically, you can unzip them using `unzip <file_name>`. From there, you will find a variety of files and directories, mostly XML files.

Every decompressed Word document contains at least two ".rels" files. These files are XML formatted with *Relationship* tags that each contain a *Target* attribute. In most circumstances, all of these *Target* attributes have a value equal to the path to some XML file within the Word document. However, *Follina* demonstrated that the *Target* attribute can also equal a web address. <u>This is a feature, not a bug.</u>

The objective of this tool is to automate the process of reading the *Target* attributes of every ".rels" file within a Word document. The results are then printed so the user can view these attribute values and decide for themself whether or not the document is trying to carry out this exploit.

## How to use the tool

You will need `git` and `python3` installed on your system. Then,

```shell
git clone https://github.com/ir1descent1/analyze_word_rels_targets.git
cd analyze_word_rels_targets
```

You can then analyze documents using:

```shell
python3 analyze.py /path/to/doc/name.docx
```

Example output:

```
*** Analyzing rels from 'example.docx' ***
Targets from _rels/.rels:
         docProps/app.xml
         docProps/core.xml
         word/document.xml

Targets from word/_rels/document.xml.rels:
         webSettings.xml
         settings.xml
         styles.xml
         theme/theme1.xml
         fontTable.xml

```

## Note on Operating Systems

This tool was designed to work on Windows and Linux, yet has not been tested on Windows. If you encounter errors on Windows, please let me know.
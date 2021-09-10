# Clash-Detection-Matrix-Automation
Automating Clash Detection through a workflow from Revit, Dynamo and ending in Navisworks using Python scripting as a BIG Data Analytics tool

## What can you do with this workflow?
1. If you would like to have information of all your Revit modelled **Elements** and **Categories** for both linked and active/open revit models, exported to an **Excel spreadsheet** in just one click, then the first part of this workflow would be just great.
2. If your company or business works with an Excel Clash Matrix template as guide for carrying out Clash detections in Navisworks, then this workflow can help you auhtomate the creation of your Clash Matrix template based on the project's disciplines, model categories and element types. This can be done in just one click
3. If you need an update of newly modelled elements in your project (for all disciplines, linked and active models) with an Excel spreadsheet capturing the changes in realtime. Then you'll benefit a lot from this.
4. If you don't want to go through the boredom involved in manually creating search sets and populating clash tolerances in Navisworks by automating the whole process, then follow through


## Getting Started
1. As an Architect, Engineer, BIM Manager/Coordinator who will find this workflow useful, i believe you have Revit and Navisworks already installed on your computer. If not you can download the 30 days free trial here; [Revit Download](https://www.autodesk.com/products/revit/free-trial), [Navisworks Download](https://www.autodesk.com/products/navisworks/free-trial)
2. Download and Install [Python](https://www.python.org/downloads/)
3. Download and install any [IDE](https://www.google.com/search?q=IDE&rlz=1C1JJTC_enNG967NG967&oq=IDE&aqs=chrome..69i57j0i271l3.2128j0j4&sourceid=chrome&ie=UTF-8) of your choice. I would recommend [Visual Studio Code](https://code.visualstudio.com/download)
4. Install the following Python modules/libraries; [Pandas](https://pandas.pydata.org/docs/getting_started/install.html), [NumPy](https://numpy.org/install/), [Openpyxl](https://openpyxl.readthedocs.io/en/stable/), [xlsxwriter](https://xlsxwriter.readthedocs.io/getting_started.html), [lmxl](https://lxml.de/installation.html);  The link describes the process. P.S: Ensure to use the **pip** installation option for all the python modules/libraries
5. You need just basic understanding of Dynamo and Revit although you should have developed your skills in Python and Navisworks to an intermediate level before you can customize the workflow to suite your business.
6. In general, anyone can utilize this workflow in their business because it's just a matter of few click to make the trick happen
7. If need be you can have a walkthrough tutorial on the Python Data Analytics libraries i utilized in the process. Some of which are; [Pandas](https://www.w3schools.com/python/pandas/default.asp), [NumPy](https://www.w3schools.com/python/numpy/default.asp), [xlsxwriter](https://xlsxwriter.readthedocs.io/index.html), 


## Introduction
Python is the programming language used in developing this workflow outside Revit. However, the whole process kickstarts in Revit during or after the BIM model authoring. 
The goal is to carry out clash detection project coordination at any phase of a BIM project. This is achieved by running through the workflow again and again when changes are made in the Revit model. Alas, it takes just few minutes when compared to the traditional method of doing it manually in Navisworks.


## Workflow
In the git repository files uploaded above, you will find folders for the whole operation. The workflow is explained in the images below and a quick video shhowing a sample workflow is also provided.
  ### Few facts to consider in this workflow;
- It works for any Revit project with other Linked models like Structural and MEP models
- The Dynamo files would come up with some **WARNINGS** but not to worry, it doesn't affect your script, the results would still come out fine
- I recommend you run the Dynamo script using **Dynamo Player**. And if you don't know how to use Dynamo Player, you can check [here](https://www.youtube.com/watch?v=R8usi9c2BVg) or [here](https://www.youtube.com/watch?v=oCDE_t6XoLI)
- Make sure to install all the python packages/modules/libraries i used in both python script
- To access all the files and documents for this workflow, check out the **Go To FIle** option here on my page in GitHub. All files are well arranged in each folders.
- Sample Revit models have been provided for you to test the workflow. It's located in the **Sample Revit Models** folder
- The Workflow How-To Video and PDF description is provided in the **Workflow Docs** folder


## Dealing with Issues
If you have any issues with the workflow, check the list of currently Open and previously reported issues for anything similar to yours. If you can't find an issue similar to yours, you can comment out your issues and it will be addressed


## What's the Future Holds
I intend coming up with a free **Revit plugin** that encapsulates all the processes involved in this clash matrix automation workflow. Work has started to achieve this. Hence, in the next few months, news about the Plugin release will be out. 

Pray for me so it won't take so much time

## Much love! :v::heart:


![Artboard 1](https://user-images.githubusercontent.com/68663705/132239256-af8d4e6d-e407-4857-b83a-0c3bb9736ccd.png)

![Artboard 2](https://user-images.githubusercontent.com/68663705/132246751-35a5fed3-41b2-46ad-97c1-867bb19ad519.png)

![Artboard 3](https://user-images.githubusercontent.com/68663705/132246778-6eee8c32-4e6f-4fb6-b23e-103adb57fed4.png)



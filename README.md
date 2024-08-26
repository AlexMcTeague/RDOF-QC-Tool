# RDOF QC Tool

During a month-long "QC push", I discussed with my coworker Brandon Trapp that we didn't have fast and easy ways to check certain data; most of our QC is done manually. Together we asked management if we could set aside time to work on a new tool to aid in performing this QC. Given the success of ![DesignShark](https://github.com/AlexMcTeague/DesignShark), they agreed, and allowed us a few hours each week to collaborate. The RDOF QC Tool is the result of our work, and made up for the time investment in spades; today it helps shave hours off of every QC, greatly improves the quality of my department's work _prior_ to submission, and allows users to check for errors that would otherwise be too time-intensive to do manually.

**Disclaimer:** The RDOF QC Tool is available in this repository, but the sample data used in the screenshots and videos below is not publicly available. I can provide a full demonstration of the RDOF QC Tool's features on request.

**Note:** The macros ran much slower while recording software was active; in reality most of the RDOF QC Tool's macros are near-instantaneous.

# Features

**File Imports:** On this page, the user selects the deliverables to be checked. They can either select each deliverable individually, or select a folder to import all 12. In addition, the user needs to import several 'forward traces' which map out the OLT.

**Errors:** Errors from the other tabs are collected on this page, so the user can read over all of them at once. This serves as an exhaustive list of the errors checked by the tool: if an error type is _not_ listed on this list, it needs to be checked manually. There are dropdowns for each error type, so the user can see exactly which piece of equipment flagged the error.

**Splice Report:** The Splice Report is checked for many types of errors, including naming and splice types. Of particular interest are "double-splices": the splice report tool is capable of detecting corrupted splices which appear in the report, but are difficult if not impossible to find by hand, given the length of each splice report.

![Splice Report QC](https://i.imgur.com/A7y0rxN.png)

**Port Continuity:** In any design, it's important to ensure that every address sends and receives signal all the way to/from the OLT. This is difficult to check by hand; you would need to do a manual trace from every activated port in an OTE/MST, or do a forward visual trace which loses accuracy, as it doesn't highlight individual addresses. The Port Continuity checker instead compares the HAF to several forward traces of the OLT, quickly verifying each addresses in a fraction of the time.

![Port Continuity QC](https://i.imgur.com/8Kbu1d0.png)

**BOMs:** This tab extracts data from the BOMs and Overall BOM, and displays it in a dashboard for easy viewing. Macro also checks both documents for errors, and compares the two to find inconsistencies.

**KMZ:** One of our design deliverables is a KMZ file, which shows all addresses, fiber equipment, etc in Google Earth. It's possible to manually check this file for mistakes, but doing so requires clicking each element to see its data. The KMZ tab of the QC tool extracts this data into multiple Excel tabs, which makes it very easy to sort and filter the data while checking for mistakes. In addition, running the KMZ macro creates a new KML file. This new file has multiple folders, representing each type of data for each class of equipment. The user can enable these folders separately to display data as visible labels, removing the requirement to click each element.

![Click to view KMZ demo video](https://i.imgur.com/wxmoIIK.mp4)

**MQMS/Prism:** The RDOF QC Tool can extract data from all deliverables in a completed project, and display it in a Userform. This userform visually matches the formatting in MQMS and Prism, allowing the user to easily fill out these web forms, or compare the data if they're already filled out.

# See also: ![DesignShark](https://github.com/AlexMcTeague/DesignShark)

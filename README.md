# Ulrich’s Web API &amp; Elsevier Scopus API in Excel

ENUG 2017 Presentation

Using Excel and VBA with APIs to Wow Your Colleagues and Patrons, presented by Hao Zeng – Yeshiva University, Annamarie Klose Hrubes - William Paterson University

Application Programming Interfaces (APIs) are a powerful way for libraries to gather data, including usage statistics, bibliographic information, and metrics. In the context of an academic library, the presenters will describe their workflow for pulling data from APIs into Excel with Visual Basic Applications (VBA). Using two examples , they will explain how and why this method provides a user-friendly solution that empowers other departments, inside and outside the library, to gather data independently. They will demonstrate their use of Scopus and Ulrich’s Web APIs.

Refer to the Ulrich's Web API documentation (https://knowledge.exlibrisgroup.com/Ulrich%27s/Product_Documentation/Configuring/Ulrichsweb_API/Ulrichsweb%3A_Using_the_Ulrichsweb_API#Ulrichsweb_API) and Elsevier API documentation (https://dev.elsevier.com/) to adjust the macros to your desired fields.

Note: An Ulrich's Web subscription is necessary to obtain an Ulrich's Web API key.

# Example: Elsevier Scopus API in Excel

## Step 1: Register at the Elsevier Develop Protal (https://dev.elsevier.com/index.html), request the Scopus API key.

## Step 2: Activate Developer tab in Excel (Please Google).

## Step 3: Create a spreadsheet:

![alt text](https://user-images.githubusercontent.com/12193996/31698201-48e19adc-b38a-11e7-9e28-a4129488a1e1.png)

## Step 4: In Developer tab, INSERT an Active X Controls Button

![alt text](https://user-images.githubusercontent.com/12193996/31698225-5d73581e-b38a-11e7-9e62-8a7045629c0f.png)

### In Developer tab, Design Mode, right click the Button, select Properties;

![alt text](https://user-images.githubusercontent.com/12193996/31698267-97177398-b38a-11e7-9e26-3045b9363b1b.png)

### Name it Scopus

![alt text](https://user-images.githubusercontent.com/12193996/31698299-da4a273c-b38a-11e7-9212-0c90fb998e73.png)

## Step 5: Set VBA Project Reference: Microsoft XML, V3.0 or V6.0
![alt text](https://user-images.githubusercontent.com/12193996/31749780-0d4df612-b44a-11e7-8ed2-04941cda9788.png)

### Insert the script between:

![alt text](https://user-images.githubusercontent.com/12193996/31698325-050dae08-b38b-11e7-8f8f-23d84a2bb113.png)

![alt text](https://user-images.githubusercontent.com/12193996/31698357-28645046-b38b-11e7-992d-a86b81372771.png)

## Step 6: Hit the button and run the script.
![alt text](https://user-images.githubusercontent.com/12193996/31749795-1c1d0480-b44a-11e7-9aee-776bc48bf7ce.png)


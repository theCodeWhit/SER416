# VBA BEGINNER TUTORIAL (LEARNIT SERIES)
>https://www.youtube.com/watch?v=G05TrN7nt6k&ab_channel=LearnitTraining

>> This is all super basic stuff, not much on scripting 
```
The Macro Recorder 13:40
Using Relative References 20:39
Recording Simple Macros 27:51
  - How to edit formatting, change phone number and SSN with recorded macros and "***-**-"&RIGHT(A1,4))
Multi-Step Macro Recording 39:25
Sort and Filter Macro Recording 45:26
Protecting and Formatting Sheets with the Macro Recorder 51:35
VBA Interface Setup 55:09
Recorder Code vs. Manual Code 1:01:11
Introduction to Editing Macros 1:12:28
Debugging Macros 1:16:53
Grammar in VBA 1:28:00
Macro Scripting Basics 1:33:04
Range 1:40:41
Selection & Color 1:47:14
Value and Clear 1:52:16
ActiveSheet, Sheets, and Name 1:54:01
CurrentRegion 1:56:40
Practice 1:58:08
```

#### VBA INTERFACE SETUP


You can use the "immediate window" to test snippets of code. To access in your VBA interface: view->Immediate Window. </br>

You can use the object browser to view all the objects in your library

#### CREATING MACROS FROM SCRATCH 
![image](https://user-images.githubusercontent.com/48422525/155425467-c4347d48-2c8e-41fe-b6ee-1de3edab6419.png)

#### RANGE
To refer to a cell (or cells): Range("C3") </br>
To select a cell (or cells): Range("C3").Select 

![image](https://user-images.githubusercontent.com/48422525/155425754-45c0b459-0f59-412e-bad8-3d674af002ea.png)

### SELECTION & COLOR 
* Range is absolute 
* Selection refers to what is already selected 
![image](https://user-images.githubusercontent.com/48422525/155426630-d1a29a55-efc6-4572-83aa-0b43ba0c5246.png)

![image](https://user-images.githubusercontent.com/48422525/155426651-322ff339-f71c-4cdf-b9a4-58940907a031.png)

![image](https://user-images.githubusercontent.com/48422525/155426756-dacab842-6bb7-41b4-8cef-98abcf3f657a.png)

![image](https://user-images.githubusercontent.com/48422525/155426783-4df921b2-b9e0-4ec3-a133-b27ca86aceac.png)

![image](https://user-images.githubusercontent.com/48422525/155426874-df33deb7-aee7-43ff-8683-64e7a056eabb.png)

#### VALUE & CLEAR

![image](https://user-images.githubusercontent.com/48422525/155427191-1b1f6dc5-5a53-4c46-81a1-8a6b2353b548.png)

#### ACTIVESHEETS, SHEETS, and NAME
To rename the sheet you have selected 
```
ActiveSheet.Name = "Portfolio4"
```
To select a sheet 

![image](https://user-images.githubusercontent.com/48422525/155427405-3e687b23-3687-4f00-9701-fbc935104d27.png)


#### CURRENT REGION 
![image](https://user-images.githubusercontent.com/48422525/155427670-e2e2d602-397a-4eaf-a727-50d0ee736522.png)



# VBA ADVANCED TUTORIAL 
> https://www.youtube.com/watch?v=MeKL_n6SiYY&ab_channel=LearnitTraining
```
Start 0:00
Variables 0:03
Variable Rules 7:59
For Next Loop Basics 11:18
For Next Loop Doubled 15:44
For Next Loop Triple 19:47
For Each Loop 22:11
Exit For Statement 23:50
Do While Loop 27:07
Do While Not Empty Loop 30:32
Do Until Loop 32:53
Do Loop UNTIL 35:11
Count and Offset 36:49
End, Address, Call Statement 48:23
Practice 55:35
Using the FIND Tool in a Macro 1:03:47
Message Boxes 1:14:50
Input Boxes 1:22:58
Code Continuation Character and vbCrLf Constant 1:27:16
If Then, ElseIF, and Else 1:33:44
Select Case 1:40:13
Multiple Variables 1:42:58
Practice 1:46:09
```
#### VARIABLES
![image](https://user-images.githubusercontent.com/48422525/155521787-c5f3eba7-8d76-412e-a793-0d00d76e4b5f.png)


![image](https://user-images.githubusercontent.com/48422525/155522171-19922afe-c24b-4e82-b2e6-c180fd4c2f77.png)

#### FOR NEXT LOOP BASICS

![image](https://user-images.githubusercontent.com/48422525/155529552-62eb7024-098d-4795-8aee-a9d702923255.png)

#### FOR NEXT LOOP DOUBLED
The trick to a doubled for loop is making sure the loop that is on the inside also ends ("Next intRow") on the inside

![image](https://user-images.githubusercontent.com/48422525/155530728-08a9fff5-1178-4972-9d7d-292732469842.png)


#### FOR NEXT LOOP TIRPLE
![image](https://user-images.githubusercontent.com/48422525/155531494-792d9f09-f2b4-4465-ad6b-2f906b9d98cc.png)


#### FOR EACH LOOP
A for each loop is a for next loop, but it operates on entire collections. In this case we have a collection of worksheets in this workbook altogether they make up the collection called ' Worksheet', so if i have something thats a collection like worksheets then I can say I want to do this particular thing to every single thing in that particular collection.

This example shows a message box for every instance of 'x' in the entire worksheet. 
> **This could be used to notify suppport staff of payment deadlines, where x would be a range of dates** 

![image](https://user-images.githubusercontent.com/48422525/155536192-4709bcef-0b5b-4276-8269-c7a58fd0e2d5.png)

#### EXIT FOR STATEMENT with if/then
The exit for statement is used to exit the for loop prior to the actual ending statement 

If you have a clause that would require you to exit the loop early, use exit for statement

![image](https://user-images.githubusercontent.com/48422525/155537358-6e97d59f-c90b-4500-87a3-042d1b3bcfc1.png)


#### DO WHILE LOOP 
Lets say we wanted to create a loop whose sole job wasnt just to do that loop a certain number of times, but maybe simply perform its task ad infinitum until some particular criteria was met, this is a "do" loop. There are a few different types of "do" loops. 

Do while loop  
![image](https://user-images.githubusercontent.com/48422525/155538976-b5365afc-3e0b-4df9-b861-7c6445e9f821.png)

#### DO WHILE CALC LOOP 
Do while not empty loop 

![image](https://user-images.githubusercontent.com/48422525/155540169-eae4cf9a-44fb-447f-9975-c642df72e491.png)

#### DO UNTIL 
Performs its action until a criteria equivocates to true. In this case, until the cell is empty 

before running macros: 
![image](https://user-images.githubusercontent.com/48422525/155541751-8bfee131-af74-442d-ad29-7c60f02c76f3.png)


After running: 
![image](https://user-images.githubusercontent.com/48422525/155541895-e7ad8f67-0f6d-467d-b4af-00d7bfd482a4.png)


#### DO LOOP UNTIL 
Instead of saying do until, we are going to say do ...loop until 

![image](https://user-images.githubusercontent.com/48422525/155542477-e71a5d68-77d2-4a31-b075-fe81181e9342.png)


### COUNT AND OFFSET
![image](https://user-images.githubusercontent.com/48422525/155547588-075fb0be-e82c-4b9e-829b-658299ff2476.png)

### ADDRESS, CALL, END
How to create a macro that travels along a particular column to find the content that is there. It will also remember exactly where it found that content so that later it can grab all that content and then be able to copy and paste it or move it to another location. 

![image](https://user-images.githubusercontent.com/48422525/155556077-51d4a97a-131f-48ed-b5f1-a1fda8c27341.png)

This can work on any of your sheets for column A.
Now that you are able to find the content, you will want to run the other macros (copy/paste/insert). 
The parent macros job is to run all the other macros. To do this, use the Call statement 


![image](https://user-images.githubusercontent.com/48422525/155556682-94802976-ba83-4d11-ad7b-70ef2c05dec5.png)

![image](https://user-images.githubusercontent.com/48422525/155556812-73229c8a-f013-47af-9a96-7a7a6b7ed88f.png)

#### REPORT GENERATOR & PRACTICE
Implementing all created macros to generate a report 

![image](https://user-images.githubusercontent.com/48422525/155576419-c315b27b-a7c6-4687-ba59-39cf8c6c7dbc.png)


### USING THE FIND TOOL 

Use a macros recorder to find any cell containing the word "portfolio" and 
add script to format it a particular way. 

We will need to put a loop in here and eventually a logic test that is able to end the loop when a particular criteria is met.

So we are going to use a do loop that runs while the cell we started on is not the cell that I have ended on 


![image](https://user-images.githubusercontent.com/48422525/155581105-2aa2f136-43bc-4b45-b3d7-e72fbf8c1f56.png)


Next you have to Call the function in your GenRep() 

### MESSAGE BOXES 


![image](https://user-images.githubusercontent.com/48422525/155581529-932f6c82-2c11-499a-b60a-81f642002e13.png)


![image](https://user-images.githubusercontent.com/48422525/155581652-e4499025-bae6-4d08-bf01-44ff312b8770.png)


![image](https://user-images.githubusercontent.com/48422525/155582060-4ea3751d-1597-4d37-b5c0-8325d9534538.png)

![image](https://user-images.githubusercontent.com/48422525/155584047-f66c2f08-afa9-472a-b458-1dcc9c93f1c9.png)

![image](https://user-images.githubusercontent.com/48422525/155584134-7a32533b-d6d5-4c62-9d16-b0496121c70a.png)

![image](https://user-images.githubusercontent.com/48422525/155584541-5a7b1794-7d62-4f37-acb9-0b63cd85e0d5.png)




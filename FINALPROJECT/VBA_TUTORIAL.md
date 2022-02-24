# VBA BEGINNER TUTORIAL 
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

#### EXIT FOR STATEMENT 


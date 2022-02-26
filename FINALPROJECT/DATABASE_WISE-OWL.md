# EXCEL VBA INTRO PART 31
> Wise Owl: https://www.youtube.com/watch?v=p_IhfObVc6A&ab_channel=WiseOwlTutorials

```
- Recap of ADO Library, Connections and recordsets
- Recordset Lock Types
- Inserting New Records
- Changing field Values
- Modifying existing records
- Looping Through a Recordset
- Deleting Records
- Transactions
- Dealing with Related Tables 
```

## CREATING A CONNECTION WITH ACCESS IN EXCEL VBA

- Creating a Connection


![image](https://user-images.githubusercontent.com/48422525/155846702-8401ff45-e36c-4415-857c-637f5e4b62de.png)

- ConnectionString for Access
  * ConnectionString sets where your database is store, what type of database it is and its security parameters
  * There is a website for finding connection strings: www.conectionstrings.com
  * It is easier to store your datastring as a constant at the top of the module 

![image](https://user-images.githubusercontent.com/48422525/155847059-4041c120-f37a-4e6e-9723-99bcd691e6cc.png)

![image](https://user-images.githubusercontent.com/48422525/155847127-058e8c9a-27c3-4361-9ab1-a12abef45466.png)
- Data Recordset 
 * Declare Data Variable (MoviesData) as recordset
 * Create an isntance of that recordset
 * Change Properties within your open connection
    * Specify which db connection your recordset will be using

![image](https://user-images.githubusercontent.com/48422525/155847881-4043d168-ba66-47a7-ba9d-61c5f67274af.png)

* Source (recordset property) 
    

![image](https://user-images.githubusercontent.com/48422525/155848112-a365ced1-cd43-462b-a661-a4f205c2895b.png)

  * Access DB view: 

![image](https://user-images.githubusercontent.com/48422525/155848364-bd5d4839-d3c5-4518-9cd5-39f60043736d.png)

* Locktype (recordset property)
  * important if you are modifying data in a database
  * avoid "read only" if you are trying to modify the data
  * adLockOptimistic controls exactly when the records get logged as changes are being made 
    * www.w3schools.com provide a more descriptive explaination of how these properties work 
    
![image](https://user-images.githubusercontent.com/48422525/155848640-1f549ecf-af50-4006-8d89-12a899fc238f.png)

* Cursor Types (recordset property)
  * This determines how dynamic your recordset will be, in otherwords, what changes it will detect
  * The most effiecent but least dynamic: OpenForwardOnly
     * You can only move forward through the recordset and you wont be able to detect changes made by users at all
  * Most dynamic but least efficent : adOpenDynamic
     * This will essentially detect all changes made by endusers but it comes at a cost in terms of performance
  * More on this at www.w3schools.com
  
![image](https://user-images.githubusercontent.com/48422525/155849622-bc3f41dc-b154-45fc-848c-c601ae902a1c.png)

* Opening and Closing a Recordset
  * you can open and close a recordset in the same way you would open/close a connection 
  * apply open method within you With statement
  * Close data outside of With statement 
  * It is important to close a recordset before you close a connection
  * You can also release the variables, such as "MoviesData" to free up the memory space (this usually already happens at "End Sub" but it is good practice 


![image](https://user-images.githubusercontent.com/48422525/155849990-4127847f-1a25-471f-9a4a-234bc41e956d.png)




## ERROR HANDLING 

We want to make sure that we close out our connection if something fails afer we have opened a connectiuon with our database. 


* Error handling for db connection (Open)


![image](https://user-images.githubusercontent.com/48422525/155850938-073e7442-b51d-4f44-b1aa-f1a00b6d3d82.png)



* Error handling for recordset (Open) 

![image](https://user-images.githubusercontent.com/48422525/155851019-08eea42e-ad7e-41e2-810d-6d9d3f96e35a.png)


* Where you modify your data

![image](https://user-images.githubusercontent.com/48422525/155851132-8016bcfa-2bcd-43c3-9f71-9e1ec66630c1.png)

* Or you can move your End With statement down 

![image](https://user-images.githubusercontent.com/48422525/155851260-0ade7aaf-e78f-463a-acf1-0e24c0a5f410.png)

## ADDING TO THE RECORDSET
* Using .AddNew
  * has a field and value property 
  * Fields refer to table column titles 
  
  ![image](https://user-images.githubusercontent.com/48422525/155855758-4fbc1a00-743e-4640-bf77-60913e742198.png)

* All together

![image](https://user-images.githubusercontent.com/48422525/155857669-8c2034cd-ec76-468f-b8ba-082a40ce8871.png)

## LOOPING OVER A RECORDSET
eof (end of file)

```
do until .EOF
...
Loop
```

## CHECKING FIELD DATA
- .MoveFirst : moves to the first element in the list
- .MoveLast : Moves to the last element in the list
- .MoveNext : next element 
```
if .Fields("FilmRuntimeMinutes").Value > 180 Then 
    .Fields("FilmRuntimeMinutes").Value = 180
    .Update
End if
.MoveNext
```

To only cancel Update if you are not at the EOF or BOF
![image](https://user-images.githubusercontent.com/48422525/155858308-e82bc191-19b2-4b88-ad19-d2178b34f63a.png)


## DELETING RECORDS

add code that allows you to undo. Before you try to modify any data you tell the db to begin a transaction [ADO terms: this is a method that is associated with a connection object rather than the record set]

* You must do 1 of two things at the end of your transaction: 
* commitTrans : commits changes
* RollbackTrans : undoes any changes since the transaction began 

You can use a conditioning statement to decide whether you want to commit or rollback. Rollback is typically used as a error handling routine in case something goes wrong with modifying data in a complex sequence of changes. 

![image](https://user-images.githubusercontent.com/48422525/155858497-272cdb4b-9e02-4470-aff7-f9e079765e15.png)






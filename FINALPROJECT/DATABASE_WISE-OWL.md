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

### CREATING A CONNECTION WITH ACCESS IN EXCEL VBA

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
 * Create and isntance of that recordset
 * Change Properties within your open connection
    * Specify which db connection your recordset will be using

![image](https://user-images.githubusercontent.com/48422525/155847881-4043d168-ba66-47a7-ba9d-61c5f67274af.png)

* Set source of recordset to a table in your access db 
    

![image](https://user-images.githubusercontent.com/48422525/155848112-a365ced1-cd43-462b-a661-a4f205c2895b.png)

* Access DB view: 

![image](https://user-images.githubusercontent.com/48422525/155848364-bd5d4839-d3c5-4518-9cd5-39f60043736d.png)






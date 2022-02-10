## USERFORM CONTROLS

Insert UserForm in VBA terminal and play around with the different tools
</br>

The TextBox tool can be used to capture input text from the user. 
Whatever you name this textbox will be how you refer to it in your VBA code. </br>
--Same applies to all tools </br>

#### COMMAND BUTTON, TEXTBOXES & LABELS
Here is an example of a simple UserForm that takes user input and places it in the label named "Username" </br>
Then dumps it into the spreadsheet (A1) once the commandButton is pressed </br>
![image](https://user-images.githubusercontent.com/48422525/153493513-bc75dfbd-480c-4128-b342-5cedb71e2201.png)



![image](https://user-images.githubusercontent.com/48422525/153492039-4e58a3ad-2055-4996-901b-8d80e3a38f2e.png)

</br>
When ran :

![image](https://user-images.githubusercontent.com/48422525/153493756-1be11739-9c26-4c1e-aa6b-fa7c88c5767e.png)

After entering Whitney into the textbox and pressing the command button, the Username changes to Whitney and Whitney is placed in A1

![image](https://user-images.githubusercontent.com/48422525/153494155-554ed34c-8cdc-445d-a73a-b040f1be2e70.png)

#### COMBOBOX (AKA DROP-DOWN MENU)
Here is a simple example of how to create a ComboBox and add options to select from  

![image](https://user-images.githubusercontent.com/48422525/153494848-a2706373-bd04-4c35-8381-0b2ffb930eab.png)


![image](https://user-images.githubusercontent.com/48422525/153490706-ad7d7956-7eae-4825-b1f5-4713e9df241e.png)
 
 Or you can add multiple items like this: 
 ~~~
With Me.ComboBox1
.AddItem "Charge"
.AddItem "Credit"
End With
~~~

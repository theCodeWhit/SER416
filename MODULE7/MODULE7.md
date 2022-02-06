# PROJECT SCHEDULE (TIME) MANAGEMENT
>12 MIN
steps involved in developing initial plan
- scope: charter -> Scope -> UML UC -> Requirements -> WBS
- Schedule: WBS -> Gantt, Network Analysis

You need to know what features you are going to have in order to make your gantt chart. 
- you will usually have an end date that is not negotiable 
- The schedule is usually the most important 

ELEMENTS OF A SCHEDULE MANAGEMENT PLAN 
its about the tools and how you plan to manage the schedule 

> a milestone is a significant event that normally has 0 duration
>   it often takes several activities 

Effort and duration 
effort might take 2 weeks but it might take 3 weeks to complete. 
effort does not help you in your schedule
you have to account for holidays and vacations, especially if it is the end of the year. 

# DEVELOPING THE SCHEDULE 
> 9 MIN

Set up initial schedule from WBS and perhaps initial cut during scope activities

The whole point is to have realistic timeframe of when the project will be complete
Things do not always start right after the previous stage is finished, sometimes there are buffers.</br>
#### IMPORTANT TOOLS AND TECHNIQUES
~~~
Gantt Charts
Critical Path Analysis 
Critical chain scheduling
PERT analysis 
~~~
#### ESTIMATING ACTIVITY DURATIONS
- Duration includes the actual amount of time worked on an activity plus elapsed time
- People doing the work should help create / confirm estimates
- A three-point estimate is a way of giving a range that includes an optimistic, most likely and pestimistic estimate
- Network and PERT -Best early for birds-eye view 
</br>
Use common practices and tools to show the schedule </br>

-  Diamonds: milestones
- Thick Lines: Duration of tasks
- Arrows: Dependencies

Dont forget to update the WBS dictionary after making your Gantt Chart 
![image](https://user-images.githubusercontent.com/48422525/152649313-f7e6ca5d-26f7-4288-bf46-f06db7bbf61a.png)

#### ADDING MILESTONES TO GANTT CHARTS
Many people like to focus on meeting miletsones, especially for large projects
- Milestones emphasize important events or accomplishments on projects 
- helps to see if it is on track or not
~~~
SMART CRITERIA FOR MILESTONES
- Specific 
- Measurable
- Assignable
- Realistic
- Time-framed
~~~
#### ADVICE FOR YOUNG PROFESSIONALS
Some people find estimating to be challenging, especially for their own work. It is very important to develop this skill .
- Practice estimating how long it takes you to do different activities and then take actual measurements 
- Define the activity in detail to help make better estimates
- if you realize that an activity estimate might not be a good one, let your team know as soon as possible so that adjustments can be made early in the project. 

![image](https://user-images.githubusercontent.com/48422525/152649611-7c30545a-b62f-46f4-ba46-c49b41006ef9.png)


# SEQUENCING ACTIVITIES
> 10 MIN
- Sequencing: process involves evaluating the reasons for dependencies and the ifferent types of dependencies
- A dependency or relationship is the sequencing of project activities
- Mandatory dependencies: inherent in the nature of the work being performed on a project, sometimes referred to as HARD LOGIC (FW design must be complete before test plans can be started)
- Discretionary dependencies: defined by the project team, sometimes referred to as preferred, preferential or soft logic an should be used with care since they may limit later scheduling  options (develop module A before going to module B; Purchase customer imaging sensor - depend on customer commit)
- External dependencies: Involve relationships between project and non-project activities (in the build vs buy if you decide on the buy option)

**Are dependencies good?** </br>
**Whant more or less of them?** </br>
The less the better </br>
**Which order do you like them?** </br>
Discretionary are the best becsause you are putting them in there, the others you have less control over. </br>

#### SEQUENCING ACTIVITIES METHODS
- Arrow diagram method (ADM): 
  - Activities are represented by Arrows
  - Nodes or circles are the starting and ending points of activities 
  - Only show finish-to-start dependencies
  - refer to the text for the step-by-step process of creating AOA diagrams 
  - not widely used 

- Precedence Diagramming Method (PDM)
  - Network diagramming technique in which boxes represent activities 
  - Most common method used

- Types of dependencies or relationships between activities 
  - Finish to start
  - Start to start:   Two things have to start at the same time 
      - [ e.g. 2 systems have to be live to start joint validation]
  - Finish to finish:  Two systems cant be integrated until they are both finished
  - start to finish: Rare but you can't hand off a project to someone until they are ready to start, aka you cannot finish it until they are ready to start it. 

Use these diagrams to get initial duration / schedule. Later can be used for critical path analysis.

![image](https://user-images.githubusercontent.com/48422525/152656945-f667344d-7e7f-44fb-bee9-54c568accbba.png)

(first one is more intuitive)

#### CRITICAL PATH METHOD (CPM)
The critical path is the path through your sequence of activities that is going to take the longest amount of time. 

![image](https://user-images.githubusercontent.com/48422525/152657885-8ec3f20a-0113-48f2-81cb-386a18124a73.png)
so which of these paths is establishing the schedule. The one that takes the longest is your schedule. You can have more than one critical path. Just because the path is critical does not mean the task is critical (growing grass). There may be more critical tasks that are not on the critical path. Examples of a critical path could be the ones that are considered high risk, have external dependencies or are prioritization by the customer, etc. 

#### CRITICAL PATH CALC EXAMPLE
![image](https://user-images.githubusercontent.com/48422525/152658019-18fddc18-22c0-4b22-9db4-0d3fe19b2071.png)
If you had a much larger project you would have a lot of problems calculating that. 
This is a brut force method and it is not ideal. 

# CRITICAL PATH METHOD/ANALYSIS
> 11 MIN
Critical Path Analysis Algorithm 

#### FINDING CRITICAL PATH - a better solution 
![image](https://user-images.githubusercontent.com/48422525/152658756-9384dec9-270c-4b85-a9a5-e070017efa3e.png)

- Each box - one task
- Duration - typically known 
- Four Corners
  - ***Typically not dates list, but offset from start; typically fix the units to hours/days/weeks*** 
  - EARLY START (ES): Earliest possible date that the task will be able to start
  - EARLY FINISH (EF): The earliest possible date that the task will be able to finish
  - LATE START (LS): The latest possible date (drop dead date) that the task will be able to start without affecting the overall project completion date. 
        - ***Drop Dead Date: last moment something cannot move without a day for day slip***
  - LATE FINISH (LF): The latest possible date that the task can finish without affecting the overall project completion date. 
      - another drop dead date

- Free float / slack 
    - Amount of time an activity can be delayed without delaying the early start of any immediately following activities
    - if in critical path -> how much that activity can be delayed without delaying the project
- Total float / slack 
  - amount of time an activity may be delayed from its early start without delaying the planned project finish date
![image](https://user-images.githubusercontent.com/48422525/152658914-7fc322e5-0db2-4c53-b8a8-3388df035ab1.png)

Example: </br> 
![image](https://user-images.githubusercontent.com/48422525/152659253-31156d35-7905-436a-acea-adf32dff92c7.png)
Homework assignment makes you take a file with this information and internalize it with data structures to do this calculation. 

# CRITICAL PATH METHOD/ANALYSIS ON PAPER
> 13 MIN


![image](https://user-images.githubusercontent.com/48422525/152660561-07ea3352-8cd1-4491-9cae-4fa16322bf78.png)


# CIRITICAL PATH METHOD/ANALYSIS VIA EXCEL
> 13 MIN
> Solved examples provided 


![image](https://user-images.githubusercontent.com/48422525/152683894-04a5cef1-20d7-48b0-a51e-62ec05b34cda.png)
</br> The grey cells are formulas </br>
You dont want to change any of the zeros 

![image](https://user-images.githubusercontent.com/48422525/152683918-88d472e5-c79f-4ec9-b13d-9b3cc29bab1b.png)


</br> Enter names and durations of all the tasks </br>
The free float is going to be the lowest ES of successors minus its early finish 
FF = ES(minimum of successors) - this.EF </br>
For A - you can delay this task by 5 units without affecting its successors. Total Float TF: how much slack you have until you increase the duration of the project. 



# CRITICAL PATH METHOD/ANALYSIS
> 5 MIN
You want to pay close attention to the tasks that have 0 slack, these are most likely on your critical path 

![image](https://user-images.githubusercontent.com/48422525/152690313-08a21d87-6acc-4e18-88e5-76ff6925c1bc.png)

</br>
next you want to focus on taks that have the least amount of slack. You can talk to your engineers and try to get some tasks off the critical path. You always have to have a critical path, they are not bad, you just want to try and minimize the critical path duration so you can be more confident in your delivery date. 

#### USING CRITICAL PATH TO SHORTEN PROJECT SCHEDULE 
- Crashing: Reduce duration of one or more tasks typically by adding more resources. Activities by obtaining the greatest amount of schedule compression for the least incremental cost.
- Fast tracking: Doing things in parallel or overlapping them, make sure tasks are sufficiently detailed 

If you need to ask for more resources be sure to have a good justification for it. 


# SCHEDULE
> 13 MIN

#### CRITICAL CHAIN SCHEDULING 
Considers limited resources when creating a project schedule and includes buffers to protect the project completion date
* Uses the Theory of Constraints (TOC): management philosphy developed by Eliyahu M. Goldratt, attempts to minimize multi-tasking when a resource works on more than one task at a time. 

#### ADDITIONAL CONCEPTS 
* Buffer: additional time to complete a task
* Murphey's Law: if something can go wrong, it will
* Parkinson's Law: Work expands to till the time allowed - try to finish it sooner 
* Project Buffer: Additional time added before the project's due date (done at a project level). You might have individual task level buffers. Be careful --You could have serious schedule bloat which wont make you competitive 
* Feeding Buffers: Additional time added before tasks on the critical path 

![image](https://user-images.githubusercontent.com/48422525/152691202-5767b4cc-fcc9-4142-a50c-e0620370004a.png)






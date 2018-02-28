'Double Click on Set Priority Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"

'Verify & Claim the Set Priority Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").WaitProperty "innerhtml", "Candidate Task.*"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]","innertext:=.*Claim").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]","innertext:=.*Claim").Click
wait(10)
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"

'Revoke Disconnect/Set Priority Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Details","xpath:=//DIV[@id='discoMenuBar']/SPAN[2]","visible:=True").FireEvent "onmousemove"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Details","xpath:=//DIV[@id='discoMenuBar']/SPAN[2]","visible:=True").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]").Click
wait(5)
'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Revoke Task","xpath:=//DIV[3]/DIV[@role='dialog']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV/DIV[@role='button' and normalize-space()='Ok']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV/DIV[@role='button' and normalize-space()='Ok']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=filterButton","xpath:=//DIV[@id='filterButton']").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=filterButton","xpath:=//DIV[@id='filterButton']").Highlight

'Assign Set Priority/Disconnect Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[2]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[2]").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Assign UserAssign","xpath:=//DIV[@id='candidate_task_form']/DIV[3]/DIV[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//DIV[@id='assignee-select']/INPUT[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//DIV[@id='assignee-select']/INPUT[1]").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//DIV[@id='assignee-select']/INPUT[1]").Set DataTable.Value("MainContName")
Wait 10
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='assignee-select']/DIV[@role='button'][1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='assignee-select']/DIV[@role='button'][1]").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV/TABLE/TBODY/TR/TD[@role='listitem']/SPAN[normalize-space()='Vutla, Prasanth']","innertext:=Vutla, Prasanth").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Assign","xpath:=//DIV[@id='assign-button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Assign","xpath:=//DIV[@id='assign-button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=filterButton","xpath:=//DIV[@id='filterButton']").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=filterButton","xpath:=//DIV[@id='filterButton']").Highlight

'Validate the Task Claimed or Not
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Details","xpath:=//DIV[@id='discoMenuBar']/SPAN[2]","visible:=True").Highlight


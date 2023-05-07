Download Link: https://assignmentchef.com/product/solved-parsing-strings-solved
<br>
<strong>Chapter 6 – Case 2</strong>

<strong>Parsing Strings</strong>




Figure 1

<img decoding="async" data-recalc-dims="1" data-src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/139.png?w=980&amp;ssl=1" class="lazyload" src="data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==">

 <noscript>

  <img decoding="async" src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/139.png?w=980&amp;ssl=1" data-recalc-dims="1">

 </noscript>

<strong>Instructions</strong>

<strong> </strong>

In this case, you will create a Visual Basic solution that manipulates strings. It will parse a string containing a list of items within a text box and put the individual items into the list box. It will build the textbox string by putting the list box items together into a single string. Parsing is based on the selected delimiter.




<strong>Step 1: Create the Project:</strong>

Create a Visual Basic Project using the project name “StringParser”.




<strong>Step 2 – Design the Form:</strong>

Design the form as shown in Figure 1.  You will need three group boxes, three radio buttons, four button controls, one text box, one list box, and two label controls.




<strong>Step 3 – Add code in the Form’s Load event to select the first radio button:</strong>

In the Form’s Load event, write code to select the Comma choice in the radio buttons group box.




<strong>Step 4 – Add code in the Parse Text to List button’s Click event to parse and load the list box:</strong>

You will need the variables shown in Table 1:




<table>

 <tbody>

  <tr>

   <td><strong>Variable name</strong></td>

   <td><strong>Type</strong></td>

   <td><strong>Use</strong></td>

  </tr>

  <tr>

   <td>delimiter</td>

   <td>String</td>

   <td>Holds the String character representing the chosen delimiter</td>

  </tr>

  <tr>

   <td>oldIndex</td>

   <td>Integer</td>

   <td>Holds the starting position in the string for the search</td>

  </tr>

  <tr>

   <td>newIndex</td>

   <td>Integer</td>

   <td>Holds the position where the delimiter was found in the string</td>

  </tr>

  <tr>

   <td>length</td>

   <td>Integer</td>

   <td>Holds the length of the input string</td>

  </tr>

  <tr>

   <td>tempString</td>

   <td>String</td>

   <td>Holds the input string from the text box</td>

  </tr>

  <tr>

   <td>tempWord</td>

   <td>String</td>

   <td>Holds the extracted word from the string</td>

  </tr>

  <tr>

   <td>advanceSize</td>

   <td>Integer</td>

   <td>Holds the number of characters to advance the pointer, to skip over the delimiter</td>

  </tr>

 </tbody>

</table>

Table 1

<u>Initialize:</u>

First, clear the list box in case there are items from a previous use.




<u>Validate and set the delimiter:</u>

Use an If/ElseIf/Else statement to validate that a delimiter has been selected. CR-LF means “carriage return – line feed”, which causes a new line to be started. You will use a built-in constant to represent this, called vbCRLF. Note that vbCRLF is two characters long, while the other delimiters are only 1 character long.




For the selected delimiter radio button, set the delimiter variable to the actual character and set the advanceSize for that delimiter, using the values in Table 2:




<table>

 <tbody>

  <tr>

   <td><strong>Selected delimiter</strong></td>

   <td><strong>Delimiter character</strong></td>

   <td><strong>Advance size</strong></td>

  </tr>

  <tr>

   <td>Comma</td>

   <td>,</td>

   <td>1</td>

  </tr>

  <tr>

   <td>CR-FL</td>

   <td>vbCRLF</td>

   <td>2</td>

  </tr>

  <tr>

   <td>Space</td>

   <td>“ “</td>

   <td>1</td>

  </tr>

 </tbody>

</table>

Table 2

Use the Exit Sub statement to leave the Click event if no radio button was selected.




<u>Parse the text box contents:</u>

Parsing a string to break out the words involves a loop and two pointer variables (oldIndex and newIndex). Both start at the beginning position, which is 0. oldIndex will always point to the current starting position for the scan (and extraction).  newIndex should be set to the position of the next delimiter. Inside the loop, do these steps:

<ol>

 <li>Scan the string from the starting position (oldIndex) until you find a delimiter. Set newIndex to the position of the delimiter.</li>

 <li>Extract the word from the starting position (oldIndex) up to but not including the delimiter position (newIndex). Use tempWord to hold the extracted word.</li>

 <li>Trim off spaces, and load the extracted word into the list box.</li>

 <li>Move the starting position (oldIndex) forward past the extracted word and past the delimiter.</li>

</ol>




<strong><em>Hints:  </em></strong>

<ol>

 <li>Use a While loop.  The condition to use will be whether oldIndex has reached the end of the input string.  You can determine this by getting the length of the input string.</li>

 <li>Use the IndexOf method to scan for the selected delimiter, and assign the results of the IndexOf method to newIndex. newIndex therefore points to the location of the next delimiter.</li>

 <li>Use the Mid function to extract the word.</li>

 <li>Remember that the delimiter size has been assigned to advanceSize.</li>

 <li>Remember that there probably is no delimiter at the very end.  So when the scan no longer finds a delimiter, you must check to see if oldIndex is at the end of the string.  If not, you should extract the remaining characters from oldIndex forward.</li>

</ol>




<strong>Step 5 – Add code in the Build Text From List button’s Click event to load the text box:</strong>




This part is simpler!  You will take the items in the list box, combine them with the delimiter, and put the final string into the text box. You will need the variables shown in Table 3:




<table>

 <tbody>

  <tr>

   <td><strong>Variable name</strong></td>

   <td><strong>Type</strong></td>

   <td><strong>Use</strong></td>

  </tr>

  <tr>

   <td>delimiter</td>

   <td>String</td>

   <td>Holds the String character representing the chosen delimiter</td>

  </tr>

  <tr>

   <td>i</td>

   <td>Integer</td>

   <td>Loop counter</td>

  </tr>

  <tr>

   <td>length</td>

   <td>Integer</td>

   <td>Holds the length of the input string</td>

  </tr>

  <tr>

   <td>tempString</td>

   <td>String</td>

   <td>Holds the input string from the text box</td>

  </tr>

  <tr>

   <td>advanceSize</td>

   <td>Integer</td>

   <td>Holds the number of characters to advance the pointer, to skip over the delimiter</td>

  </tr>

 </tbody>

</table>

Table 3

<u>Initialize:</u>

First, clear the text box of items from a previous use.




<u>Validate and set the delimiter:</u>

This code will be identical to the code used in the Parse Text to List button’s click event to validate the delimiter choice.




<u>Load the text box from the list box:</u>

You will need a loop to iterate through all of the items in the list box. For each item in the list, concatenate its value with the tempString variable.  If it is not the last item in the list, concatenate the sleeted delimiter also. After all items have been concatenated into tempString, assign it to the text box.




<strong><em>Hints:  </em></strong>

<ol>

 <li>You can use a For loop here, because you know the number of items in the list (the Items.Count property provides the count of items in the list box).</li>

 <li>You can test to see if the loop counter has reached the last item by comparing it with the Items.Count property to determine if you need to add a delimiter for this item.</li>

</ol>




<strong>Step 6 – Finish up:</strong>




Be sure to add the code for the Clear button and the Exit button.




<strong>Step 7:  Save and run</strong>

Save all files, then start the application.  Test the program using various delimiters.  Be sure to enter the data in the textbox using the delimiter you have selected.  Figure 2 shows a sample run using commas in the text box. <img decoding="async" data-recalc-dims="1" data-src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/242.png?w=980&amp;ssl=1" class="lazyload" src="data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==">

 <noscript>

  <img decoding="async" src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/242.png?w=980&amp;ssl=1" data-recalc-dims="1">

 </noscript>




Figure 2




Then change the delimiter and use the second button to load the textbox from the list box. Figure 3 shows the screen after changing the delimiter to CR-LF, and then clicking the Build Text from List button.




<strong><img decoding="async" data-recalc-dims="1" data-src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/439.png?w=980&amp;ssl=1" class="lazyload" src="data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==">

  <noscript>

   <img decoding="async" src="https://i0.wp.com/www.ankitcodinghub.com/wp-content/uploads/2017/05/439.png?w=980&amp;ssl=1" data-recalc-dims="1">

  </noscript> </strong>

<strong>Note</strong>:  There is another way to parse a string, but you need to use an array. The String method Split will split a string into its component pieces and place the pieces into an array.
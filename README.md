# WordVBAPayload
Create Word VBA payload that self-destruction at runtime, it only executes once !

This PoC show how to create and embed a binary payload, in this example reverse shell coded in C. There is many methods of embedding or fetch the code we want to execute, strings, arrays, hidden in objects or downloaded from internet, they all gets 
flagged by AV/EDR, in this example I use merge of files, like:

copy /B Document.doc+payload.txt NewDocumentWithBinaryPayload.doc (don't use .docx) the payload.txt is our binary payload, open payload.txt in a HEX editor, HxD is a greate tool.

![Step1](https://github.com/mobdk/WordVBAPayload/blob/master/step1.PNG)

Now insert some hex string that is unique, something we can serach for, to find our offset, I use f181d8, this is the beginning of 
our payload inside Word document. Save changes and copy all the HEX values, paste in NotePad++ and remove all spaces


![Step2](https://github.com/mobdk/WordVBAPayload/blob/master/step2.PNG)


Before we move on, we have to make it harder for AV/EDR to analyse our binary data, I recommend inserting fake data ever second line, so I split it up in 100 chars and insert fake data, like this:

In Notepad++ serach for: ^.{100} replace with: $0\r\n


![Step3](https://github.com/mobdk/WordVBAPayload/blob/master/step3.PNG)

every second line is fake data, starting with a1, this in not something yoy have to do, but it makes it harder to analyse our 
embedded payload. The last thing I do is adding the value 0 at EOF, so every line is 100 hex chars long, it makes it simpler
to extract from VBA code, like this: (remember to remove all line breaks)


![Step4](https://github.com/mobdk/WordVBAPayload/blob/master/step4.PNG)



Now we can serach for our offset = f181d8 and extract binary data, I have 1135 lines divided by 2 and each line is 100 divided by 2:

```
For y = 1 To 567  <--- 1135/2
  For i = 1 To 50  <--- 100/2
      Put #File2, p, CByte(buffer(pos))  <--- write payload to disk
      pos = pos + 1
      p = p + 1
    Next i
    pos = pos + 50  <--- step 50
Next y
```



 

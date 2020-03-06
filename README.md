# WordVBAPayload
Create Word VBA payload that self-destruction at runtime

This PoC show how to create and embed a binary payload, in this example reverse shell coded in C. There is many methods of embedding or fetch the code we want to execute, strings, arrays, hidden in objects or downloaded from internet, they all gets 
flagged by AV/EDR, in this example I use merge of files, like:

copy /B Document.doc+payload.txt NewDocumentWithBinaryPayload.doc (don't use .docx) the payload.txt is our binary payload, open payload.txt in a HEX editor, HxD is a greate tool.


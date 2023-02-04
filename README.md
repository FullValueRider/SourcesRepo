# twLib
 A twinBasicLibrary of useful tools I created to assist in solving Advent of Code problems

## Danger Will Robinson Danger!!

Please be aware that this library is a work in porgress and consequently contains an immense amount of cruft and abandoned thinking.

The realy useful bits for me are

• Meta  - a class to provide lots of meta data on a variable
```
Dim myMeta as Meta = Meta.Deb(myVariable)
if myMeta.isListarray then
....etc
```
• IterItems - Enumerate anything without resorting to behid the scenes jiggery pokery

```
'Typical template

 Dim myItems as Iteritems = Iteritems(myVar)
 Do
     dim myItem as variant = myItems.Item(0)   ' Gets the current item.
     myItem = myItems.Item(-3)                 ' we can do relative addressing
     dim myIndex as long = myItems.Index(0)    ' get the offset from the first index value - also allows relative addressing
     Dim myKey as Variant = myItems.Key(0)     ' get the actual index of the current item - also allows relative addressing
 Loop While myItems.MoveNext
 ```
• seq - an arraylist/collection on steroids (i.e. lots and lots of extra functionality) 

• Hkvp - an extended version of Olaf Scmidts cHashD dictionary (Hashed Key Value Pairs - geddit <groan> 

• Self factoryclasses  - all calsses use Deb as a self factory method with constructor parameters 

• Fluent Api for the seq and Hkvp classes 

• 'First classish' functions 


## Other Examples
```
Dim mySeq as seq = seq.deb(32)                 ' create a sequence with an inital capacity of 32 entries (1 to 32)
Dim mySeq as Seq = seq.deb("Hello")            ' creat a sequence of 5 characters from string "Hello"

mySet = mySeq.InBoth(myOtherSeq)               ' returns a new sequence containing items that only appear in both sequences.
Dim myHistogram as Hkvp = mySeq.freq           ' return a count dictionary of unique Items vs number of ocurrences
Dim myReverse as seq = mySeq.Sort.Reverse
Dim myLongs as seq = mySeq.mapIt(mpConvert(ToLong))   ' apply the .execmap method in class mpConvert to each item in mySeq
```


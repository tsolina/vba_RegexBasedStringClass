# vba_RegexBasedStringClass
Custom string class powered by Regex to search, replace, split and more 

Intention is to simplify use of strings with addition of string class which works similar to strings from fully
object oriented languages.

String Class:
After referencing this library to your project, its possible to use it like:
Example:
    With stringCreate("  hel8165lo world445454!").TrimLeft.ReplaceS("/\d/g", vbNullString).ToUpperCase
        Debug.Print .Value, .Length
    End With
Will print: 
    HELLO WORLD!   12

explanation:
    it will remove leading white spaces
    it will remove all numbers
    change string to upper case 
    and return the value of new string

Regex class:  
regex syntax is declared similar to javascript, as "/" + regex pattern + "/" + optional modifiers(gmi)
declaration with two strings where first string is string literal and second are modifiers is supported as well
Example:
    Debug.Print regexCreate("ha", "gi").Replace("hallo hAllo hallo", "He")
    Debug.Print regexCreate("/ha/gi").Replace("hallo hAllo hallo", "He")
both will print
    Hello Hello Hello


Additional functionalities are there to explore!

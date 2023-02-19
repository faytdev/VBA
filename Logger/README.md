# Logger for VBA

## Calling Log Method
___
The Log method is the main logging method. It can be used to log all levels. You can also call the individual LogLEVEL subs. See Public Members section.

Example below show using the Log method with all three log levels. It also shows logging the message directly with LogInfo. Then Writes the log to the immediate window and flushes the buffer.

    Sub Log()
        Dim Message As String
        Message = "Message To Log"

        Logger.Log Message 'Default level is Info
        Logger.Log Message, LogLevelDebug
        Logger.Log Message, LogLevelError
        
        Logger.LogInfo Message

        Logger.WritToConsole True
    End Sub

___
## Public Members
___
### Enums
___
* logLevel
    * LogLevelInfo = 0
    * LogLevelDebug = 1
    * LogLevelError = 2
___
### Properties
___
* Pattern
    * Read & Write
    * Sets the Start of Log line Pattern. Default is = "< " & Format$(Date + Time, "mm/dd/yyyy@hh:mm:ss") & " > ::#::  "
* Symbol
    * Read & Write
    * Sets the Symbol to replace in Pattern String. This is used to set the position of the LogLevel value. If Symbol is not in Pattern then LogLevel is appended to the end of pattern. 
* BufferArray
    * Read Only
    * Returns the current Buffer as a Variant as a one dimintion array. 
___
### Methods
___
* Flush
    * Errases Buffer and sets Buffer = 0
* FlushPatternAndSymbol
    * Sets Pattern and Sybmol = to vbNullString
* Log
    * Required msg As String
    * Optional lvl As logLevel = logLevel.LogLevelInfo
    * Use to add message to Buffer at any log level
* LogInfo
    * Required msg As String
    * Logs Info level to Buffer
* LogDebug
    * Required msg As String
    * Logs Debug level to Buffer
* LogError
    * Required msg As String
    * Logs Error level to Buffer
* WriteToConsole
    * Optional FlushBuffer As Boolean = False
    * Optional FlushPatternSymbol As Boolean = False
    * Writes current Buffer to immediate window.
    * If FlushBuffer = True then Buffer is erased
    * If FlushPatterSymbol = True then Pattern and Symbol = vbNullString
* WriteToFile
    * Required Filepath As String
    * Optional FlushBuffer As Boolean = False
    * Optional FlushPatternSymbol As Boolean = False
    * Writes current Buffer to file. If File does not exist file is create. If file exists, the file is appened.
    * If FlushBuffer = True then Buffer is erased
    * If FlushPatterSymbol = True then Pattern and Symbol = vbNullString
* WriteToRange
    * Required Target As Range
    * Optional FlushBuffer As Boolean = False
    * Optional FlushPatternSymbol As Boolean = False
    * Writes current Buffer to range. If more than one cell is selected, range is set to the to left most cell. 
    * If FlushBuffer = True then Buffer is erased
    * If FlushPatterSymbol = True then Pattern and Symbol = vbNullString
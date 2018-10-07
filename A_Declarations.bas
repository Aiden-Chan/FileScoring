Attribute VB_Name = "A_Declarations"
Option Explicit

Public Const NumFreq As Integer = 10

Public NumFiles As Integer      'Number of text files analysed
Public NameFile() As String     'Name of each text file
Public ContentFile() As String  'Content of each text file as a string
Public FolderPath As String     'Path of folder containing text files
Public Words() As String        'Array of all words in current text file
Public NumWords() As Integer    'Total number of words in each text file
Public FileType As String       'File type, e.g. .txt or .xml
Public PDF As Boolean           'Export results as PDF?

# LotusScript-CSV-Parser
* The file CSV_Parser.lss contains a parser library for CSV (Comma-Separated Values) text.
* The original code was made by Julian M Bucknall and published at http://www.boyet.com/articles/csvparser.html (thanks, Julian!)

Author notes:

_**Important:** The Lotus Script MultiByte Characterset (LSMBC) is used for storing Strings in memory. When the CSV data contains UTF codepoints that cannot be represented in LSMBC they turn into Nulls. This is obvious when working with some of the newest emojis (as of 2020). The change of data integrety maybe be unacceptable, and there is no workaround of this (not on Domino 9 or earlier)._
* The code was ported to Lotus Script by Sam Sirry, with a few modifications.
* The SignalProgress method was added to the Consumer class and it helps indicate to the caller how much of the CSV data has been processed.
* The Continue variable has been made available in the events called in the Consumer. It is True by default. If the Consumer sets it to False this will cause the Parser to stop processing and return to the caller.
* Since that Lotus Script is notoriously slow walking over long strings, the BigStringWalker class was implemeted in the CSVParser. The BigStringWalker library can be found here: https://github.com/SamSirry/LotusScript-BigStringWalker



Sample code:
```vb
'The code in SampleCsvConsumer.lss contains a sample CSVConsumer class.

Dim Consumer As New CsvConsumer(db)
Dim Parser As New CSVParser
Dim Tokenizer As New CSVCharTokenizer(CSVString)

Call Parser.Parse(Tokenizer, Consumer)
```

Keywords & keyphrases:
* Import CSV using LotusScript
* Import CSV file into Lotus Notes database
* Lotus Script processing CSV data
* Import Data from a CSV File using LotusScript

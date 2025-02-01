# TextJaccardIndexWords
Calculate word based Jaccard similarity index between texts in cells in two columns in Excel.

The index is calculated on words as split from the text using " " as separator. This means words
ending in different punctuation, or different capitalization, in brackets, etc. are counted as distinct.

The macro assumes the original text is in column D, the revised text in column E, 
and the Jaccard Index (%) will be added to column F of the row.

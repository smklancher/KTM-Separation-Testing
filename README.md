# KTM Separation Testing

## Normal Product Functionality

### Separation Benchmark

The best way to test separation in Project Builder is to use a golden set of correctly separated and classified documents and then use the Separation Benchmark functionality.  It makes it easy to see expected and actual split points.

One thing it does not do is actually let you work with the resulting separated documents in a test set.  A reason you might want to do this is so that testing in Project Builder can match runtime behavior as closely as possible.

### Test Separation

There is another way to test separation, but it takes some extra work.  With a test set in hierarchy view (which only works on the default subset):

- Manually merge a batch of documents into one document: Select all documents, right click, Edit > Merge.
- Test Separation with F9 or on the ribbon menu Process > Separate. This opens the Document Separation Results window, which shows the resulting separation and classification of the batch of pages.
- If you want to actually separate the documents in your test set, at this point you would need to note the results (screenshots?), so that you can manually split at the same points.

## Separation Testing Script

This script makes it easier.  Calling Separation_MergeDocsAndSeparate (ideally from the KTM Dev Menu) will do the same Test Separation approach without the manual steps.  The result is that this can be run on a folder of documents and they will be merged together, then split and classified as they would at runtime.  you can then test on the documents further (Extract, Validate) as needed.

### TDS Test Project

This project has a small TDS separation model trained with a small set of random example mortgage PDFs from around the internet.  As an additional example, the function ExportDocAsTiff shows how to convert PDFs to TIFF in script (optionally multipage, optionally bitonal).  This is useful because you cannot test separation on PDFs and this makes it easy to convert a folder of PDFs to TIFFs with the same names.
# KTM Separation Testing

This script makes it easier to test separation projects in Project Builder at design time in a way that more closely matches KTM Server at runtime.  Because KTA does not provide batch/folder level events to test in Transformation Designer, this script can only work correctly in KTM Project Builder.

## Normal Product Functionality

### Separation Benchmark

The best way to test separation in Project Builder is to use a golden set of correctly separated and classified documents and then use the Separation Benchmark functionality.  It makes it easy to see expected and actual split points.

One thing it does not do is actually let you work with the resulting separated documents in a test set.  A reason you might want to do this is so that testing in Project Builder can match runtime behavior as closely as possible.

### Test Separation

There is another way to test separation, but it takes some extra work.  With a test set in hierarchy view (which only works on the default subset):

- Manually merge a batch of documents into one document: Select all documents, right click, Edit > Merge.
- Test Separation with F9 or on the ribbon menu Process > Separate. This opens the Document Separation Results window, which shows the resulting separation and classification of the batch of pages.
- If you want to actually separate the documents in your test set, at this point you would need to note the results (by writing them down or taking a screenshot), so that you can manually split at the same points after closing the window.

## Separation Testing Script

To avoid the manual steps, call the included function Separation_MergeDocsAndSeparate merge all of the documents together, run Separation (classifiers and marking splits), then actually split the documents and classify them to mirror what happens at runtime.  They can then be tested further (Extract, Validate) as needed.

Ideally Separation_MergeDocsAndSeparate is called from the [KTM Dev Menu](https://github.com/smklancher/KTM-Dev-Menu) included in the project.  

### TDS Test Project

This project has a small TDS separation model trained with a small set of random example mortgage PDFs from around the internet.  As an additional example, the function ExportDocAsTiff shows how to convert PDFs to TIFF in script (optionally multipage, optionally bitonal).  This is useful because you cannot test separation on PDFs and this makes it easy to convert a folder of PDFs to TIFFs with the same names.
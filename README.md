# PHP_XLSXWriter Modified by Denuu
This is a fork of the fantastically lightweight and quick PHP_XLSXWriter by @mk-j. It's the remedy for constantly running out of memory while writing XLSX spreadsheets with PHP.
I did miss a few features from the previous library, so they've been implemented. Cross-compatible custom widths, title rows, autofilter, and freeze panes have been added.

### Original Documentation
https://github.com/mk-j/PHP_XLSXWriter

## (De)New Features
### Write Sheet Title
Sometimes (always) we want to have a title row prior to the data columns, this modification allows for just that.
Instead of passing page-wide column options into `writeSheetHeader()`, pass them into `writeSheetTitle()` as below:
```
$title_options = [
    'font-size' 	=> 20,
    'font-style' 	=> 'bold',
    'valign' 		=> 'center',
    'widths' 		=> [50, 80, 20, 40, 40, 25, 50, 40, 40, 40, 30],    			
    'autofilter'	=> true,
    'freeze_pane'	=> [2, 11]
];
$writer = new XLSXWriter();
$writer->setAuthor('author');
$writer->writeSheetTitle('Spreadsheet_name', 'title-text', [int] # of columns to span, $title_options);
$writer->writeSheetHeader(...);
```

### AutoFilter
Yes, that works now - with or without the use of the aforementioned title row. Simply pass it as a true option into the `writeSheetHeader()` options, or `writeSheetTitle()` options if you're using that:
```
$title_options = [
    ...,
    'autofilter' => true
];
```

### Freeze Panes
I've added this as well, it's really nice to have the title and column headers frozen as you scroll through hundreds of thousands of lines of data. Simple pass the range of cells you'd like to freeze as a range array `[y, x]` where `y` is the number of rows, and `x` is the number of columns to be frozen (from origin, base 1).
```
$title_options = [
    ...,
    'freeze_pane' => [2, 11]
];
```
The above will freeze the title row and the column headers, which looks good and works well.

### X-compatible Widths
The original library supports custom column widths, but the widths only function with Excel on Windows. I've modified widths to function on Mac Excel, and third party spreadsheets such as whatever Dave would use on UbuntuX. Nothing needs be done to enable this, simply pass an array of column widths into `writeSheetTitle()` if it's being used, otherwise into `writeSheetHeader()`:
```
$title_options = [
    ...,
    'widths' => [50, 80, 20, 40, 40, 25, 50, 40, 40, 40, 30],
    ...
];
$writer->writeSheetTitle('Spreadsheet_name', 'title-text', [int] # of columns to span, $title_options);
```

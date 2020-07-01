# TYPO3 extension quotation_tt_products

## What is does

This extension gives you the possibility to export the current basket view into an Excel file. This can be used to make an automatically generated quotation.

This extension is used by tt_products.


## Prerequisites

You must create the file fileadmin/data/quotation/count.txt manually and make it writable.


## How is the Excel file generation started?

Enter these lines into the Shop Template inside of the HTML form of the subpart ###BASKET_TEMPLATE###.

### Example

<!-- ###BASKET_TEMPLATE### begin -->

...

<form method="post" action="###FORM_URL###" name="warenkorbform">

...
<script>
function anexport() {
 document.warenkorbform.action = "###FORM_URL###&amp;eID=export_excel";
 document.warenkorbform.submit();
 document.warenkorbform.action = "###FORM_URL###";
}
</script>

<input type="button" name="ex" value="Export as Quotation (XLS)" onclick="anexport();">

<input type="submit" name="products_info" value="###BASKET_CHECKOUT###" onclick="this.form.action='###FORM_URL_INFO###';">

...
</form>
...

<!-- ###BASKET_TEMPLATE### end -->






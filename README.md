CakePHPExcel
============

A plugin to generate Excel files with CakePHP. Uses CakePHP view files to generate them (with PHPExcel_IOFactory class from PHPExcel).

Requirements
------------
* PHP 5.2.8
* CakePHP 2.1+
* [PHPExcel](https://github.com/phpoffice/phpexcel)
* Composer

Installation
------------

Add to your composer.json file (maraya/cake-php-excel)

```
"require": {
	"maraya/cake-php-excel": "1.1.*"
},
"config": {
	"vendor-dir": "Vendor/"
},
"extra": {
    "installer-paths": {
        "Plugin/CakePHPExcel": ["maraya/cake-php-excel"]
    }
}

```

And run

```
composer update
```

Usage
-----

In `app/Config/bootstrap.php` add:

```
CakePlugin::load('CakePHPExcel', 
	array(
		'routes' => true
	)
);
```

Add the RequestHandler component to AppController, and map Excel extensions to the CakePHPExcel plugin
```
'RequestHandler' => array(
	'viewClassMap' => array(
		'xls' => 'CakeExcel.Excel',
		'xlsx' => 'CakeExcel.Excel'
	)
),
```

Create `Layouts/xls/default.ctp` (in this example the charset is UTF-8)
```
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
	<?php echo $this->fetch('content'); ?>
</body>
</html>
```

Place the view templates in a `xls` subdir, for example: `app/View/Reports/xls/clients.ctp`

```
<h1>Title</h1>
<table width="100%" border="1">
	<tr>
		<td>col1</td>
		<td>col2</td>
		<td>col3</td>
	</tr>
	<tr>
		<td>val1</td>
		<td>val2</td>
		<td>val3</td>
	</tr>
</table>
```

And in your controller:

```
class ReportsController extends AppController {
    public function clients() {
		$this->excelConfig =  array(
			'filename' => 'clients.xlsx'
		);
	}
}
```

Call the URL

```
http://example.com/reports/clients.xlsx
```

If you want to download Excel5 format, change the URL extension from xlsx to xls.

Inside your view file you can write HTML code. Please see the [PHPExcel](https://github.com/PHPOffice/PHPExcel) documentation for a guide on how to use PHPExcel.

# Create-Microsoft-Word-File-in-Laravel
Create and Print Microsoft Word File, ODF, HTML File in Laravel
<p align="center"><img src="https://laravel.com/assets/img/components/logo-laravel.svg"></p>

<p align="center">
<a href="https://travis-ci.org/laravel/framework"><img src="https://travis-ci.org/laravel/framework.svg" alt="Build Status"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://poser.pugx.org/laravel/framework/d/total.svg" alt="Total Downloads"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://poser.pugx.org/laravel/framework/v/stable.svg" alt="Latest Stable Version"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://poser.pugx.org/laravel/framework/license.svg" alt="License"></a>

</p>

## Note
 - `composer require phpoffice/phpword`

  

If you are saving a document as an `Word File`. Edit in `DocumentController.php`

```php
    public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/python.png",array('width'=>'300','padding' => '20px'));  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save('zinmyosweSE.docx');
        return response()->download(public_path('zinmyosweSE.docx'));
    }
```
If you are saving a document as an `ODF File`. Edit in `DocumentController.php`
```php
public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/Krunal.jpg");  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
        $objWriter->save('Appdividend.odt');
        return response()->download(public_path('Appdividend.odt'));
    }
```
If you are saving a document as an `HTML File`. Edit in `DocumentController.php`
```php
public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/Krunal.jpg");  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
        $objWriter->save('Appdividend.html');
        return response()->download(public_path('Appdividend.html'));
    }
```

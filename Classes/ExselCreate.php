<?
require_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include.php");
require_once __DIR__ . '/PHPExcel.php';
require_once __DIR__ . '/PHPExcel/Writer/Excel2007.php';
\Bitrix\Main\Loader::includeModule('iblock');


class ExselCreate
{
    private $iblockId = 9;
    private $arrTitle = [
        'SECTION' => "Категория",
        'NAME' => 'Название',
        'ID' => 'Идентификатор',
        'DESCRIPTION' => 'Описание',
        'PRICE' => 'Цена',
        'PHOTO' => 'Фото',
        'POPULAR' => 'Популярный товар',
        'STOCK' => 'В наличии'
    ];
    private $alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
    private $sheet = '';
    private $xls = '';
    public $arSections = [];
    public $arItems = [];

    function __construct() {
        $this->xls = new PHPExcel();

        $this->xls->setActiveSheetIndex(0);
        $this->sheet = $this->xls->getActiveSheet();
    }

    public function creaeteFile($title)
    {

        $this->sheet->setTitle($title);
        // Формат
        $this->sheet->getPageSetup()->SetPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

        // Ориентация
        // ORIENTATION_PORTRAIT — книжная
        // ORIENTATION_LANDSCAPE — альбомная
        $this->sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
        // Поля
        $this->sheet->getPageMargins()->setTop(1);
        $this->sheet->getPageMargins()->setRight(0.75);
        $this->sheet->getPageMargins()->setLeft(0.75);
        $this->sheet->getPageMargins()->setBottom(1);

        $this->sheet->getDefaultStyle()->getFont()->setName('Times New Roman');
        $this->sheet->getDefaultStyle()->getFont()->setSize(10);

    }
    public function createTitleRow($arr = []) {
        $count = 0;
        foreach ($arr as $k => $title) {
            $coordinate = $this->alphabet[$count] . '1';
            $this->sheet->setCellValue($coordinate, $title);
            $this->sheet->getStyle($coordinate)->getFont()->setBold(true);
            if($this->alphabet[$count] == 'A') {
                $this->sheet->getColumnDimension("A")->setWidth(50);
            } else {
                $this->sheet->getColumnDimension($this->alphabet[$count])->setAutoSize(true);
            }
            $this->sheet->getStyle($coordinate)->getAlignment()->setWrapText(true);
            $count++;
        }
    }
    public function fillCel($cel, $value) {
        $this->sheet->setCellValue($cel, $value);
        $this->sheet->getStyle($cel)->getAlignment()->setWrapText(true);
        $this->sheet->getStyle($cel)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
    }

    public function saveFile($name) {
        $objWriter = new PHPExcel_Writer_Excel5($this->xls);
        $objWriter->save($_SERVER['DOCUMENT_ROOT'] . '/excel/'. $name . '.xls');
        return '/excel/'. $name . '.xls';
    }

    public function getItem() {
        $this->getSection();

        $arEl = CIBlockElement::GetList([], ["IBLOCK_ID" => $this->iblockId, 'ACTIVE' => 'Y'], false, false,
            ['ID', "IBLOCK_ID", "IBLOCK_SECTION_ID", "NAME",
                'PROPERTY_42','ELEMENT_META_DESCRIPTION',
                'PROPERTY_253', 'PROPERTY_JUSTIFICATION' ]);
        $count = 1;
        while ($arResult = $arEl->GetNext(true, false)) {
            $this->arItems[$count]['SECTION'] = $this->arSections[$arResult['IBLOCK_SECTION_ID']];
            $this->arItems[$count]['NAME'] = $arResult['NAME'];
            $this->arItems[$count]['PRICE'] = $arResult['PROPERTY_42_VALUE'];
            $this->arItems[$count]['DESCRIPTION'] = $arResult['PROPERTY_JUSTIFICATION_VALUE'];
            $this->arItems[$count]['ID'] = $arResult['ID'];
            $this->arItems[$count]['PHOTO'] = '';
            $this->arItems[$count]['POPULAR'] = '';
            $this->arItems[$count]['STOCK'] = 'Да';
            $count++;
        }
        return $this->arItems;

    }
    public function getSection() {
        $arResult['sections'] = [];
        $count = 1;
        $arSec = CIBlockSection::GetList([], ["IBLOCK_ID" => 9, 'ACTIVE' => 'Y'], false, ['ID', "IBLOCK_ID", "IBLOCK_SECTION_ID", "NAME", ]);
        while ($arRes = $arSec->GetNext(true, false)) {
            $this->arSections[$arRes['ID']] = $arRes['NAME'];
        }
        return $this->arSections;
    }
    public function fileCompletion($title, $name) {
        $res = $this->getItem();

        $this->creaeteFile($title);
        $this->createTitleRow($this->arrTitle);

        foreach ($res as $key => $row) {
            $count = 0;
            foreach ($this->arrTitle as $code => $value) {

                $cell = $this->alphabet[$count] . ($key + 1);

                $this->fillCel($cell , $row[$code]);
                $count++;
            }
        }
        $res = $this->saveFile($name);
        return $res;
    }
}



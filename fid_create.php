<?
require_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include.php");
require_once (__DIR__ . '/Classes/ExselCreate.php');

$file = new ExselCreate;

if($res = $file->fileCompletion('Фид Яндекс.Бизнес', 'fid_yandex')) {
    echo '<h2>Файл с генерирован</h2>';
    echo '<a href="' . $res . '">Нажмите чтобы скачать файл<a>';
}
else {
    echo 'Не известная ошибка';
}


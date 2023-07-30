<?php

header('Location: index.html');

require_once 'vendor/autoload.php';

//Подключаем библиотеку PHP Mailer

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;

$mail = new PHPMailer(true);

//Настраиваем почтовый сервер

$mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
$mail->isSMTP();                                            //Send using SMTP
$mail->Host       = 'smtp.yandex.ru';                     //Set the SMTP server to send through
$mail->SMTPAuth   = true;                                   //Enable SMTP authentication
$mail->Username   = 'IVashurin';                     //SMTP username
$mail->Password   = 'xkimljhgyywpotam';                               //SMTP password
$mail->SMTPSecure = PHPMailer::ENCRYPTION_SMTPS;            //Enable implicit TLS encryption
$mail->Port       = 465;     

//Подключаем шаблон документа через PHPWord

$TemplateCheckList = new \PhpOffice\PhpWord\TemplateProcessor('CheckingResultKIAflex.docx');
$TemplateOrderSK = new \PhpOffice\PhpWord\TemplateProcessor('skoringOrderSK.docx');

//Загружаем данные формы


$surname = $_POST['input-surname'];
$firstname = $_POST['input-firstname'];
$fathername = $_POST['input-fathername'];
$birthday = $_POST['input-birthday'];
$phonenumber = $_POST['input-phonenumber'];
$pasportSeries = $_POST['input-series'];
$pasportNum = $_POST['input-pasportNum'];
$issueDatePas = $_POST['input-issueDatePas'];
$pasportDivisionCode = $_POST['input-divisionCode'];
$DDseries = $_POST['input-DDseries'];
$ddNum = $_POST['input-ddNum'];
$issueDateDD = $_POST['input-issueDateDD'];
$deadlineDateDD = $_POST['input-deadlineDate'];
$startDrivingYear = $_POST['input-startDrivingYear'];
$homeRegion = $_POST['input-region'];


$fullname = $surname.' '.$firstname.' '.$fathername;
$setData = date('d.m.Y');
$fullPasNum = $pasportSeries.$pasportNum;
$fullDDnum = $DDseries.$ddNum;

//Отправка запросов и заполнение чек-листа

$TemplateCheckList->setValue('Surname', $surname);
$TemplateCheckList->setValue('firstname', $firstname);
$TemplateCheckList->setValue('fathername', $fathername);

//$curl = curl_init();

/////МВД

$curl = curl_init();
curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api-cloud.ru/api/mvd.php?type=chekpassport&seria={$pasportSeries}&nomer={$pasportNum}&token=#############",//Токен можно получить тут https://api-cloud.ru/mvd
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'GET',
));

$mvd = json_decode(curl_exec($curl), true);

if(array_key_exists('status', $mvd)){
  $TemplateCheckList->setValue('checkingMVD', $mvd['info']);
} elseif (array_key_exists('error', $mvd)){
  $er = 'Ошибка '.$mvd['error'].': '.$mvd['message'];
  $TemplateCheckList->setValue('checkingMVD', $er);
} else {
  $TemplateCheckList->setValue('checkingMVD', 'Неизвестный формат ответа');
};


/////ГИБДД


curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api-cloud.ru/api/gibdd.php?type=driver&serianomer={$fullDDnum}&date={$issueDateDD}&token=#############",//Токен можно получить тут https://api-cloud.ru/mvd
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'GET',
));

$gibddArr = json_decode(curl_exec($curl), true);
//print_r($gibddArr);


if(array_key_exists('doc', $gibddArr)){
  $gibddResult = implode('; ', $gibddArr['doc']);
  $TemplateCheckList->setValue('ChekingDD', $gibddResult);
  
} elseif (array_key_exists('messege', $gibddArr)){
  $TemplateCheckList->setValue('ChekingDD', $gibddArr['messege']);
} elseif (array_key_exists('error', $gibddArr)){
  $er = 'Ошибка '.$gibddArr['error'].': '.$gibddArr['message'];
  $TemplateCheckList->setValue('ChekingDD', $er);
} else {
  $TemplateCheckList->setValue('ChekingDD', 'Неизвестный формат ответа');
};



/////РСА


curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api-cloud.ru/api/rsa.php?type=kbm&surname='.urlencode($surname)'&name='.urlencode($firstname)'&patronymic='.urlencode($fathername)'&birthday={$birthday}&driverDocSeries='.urlencode($DDseries)'&driverDocNumber='.urlencode($ddNum)'&token=#############",//Токен можно получить тут https://api-cloud.ru/mvd
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'GET',
));


$rsa = json_decode(curl_exec($curl), true);
echo $rsa['kbmValue'];

if (array_key_exists('kbmValue', $rsa)) {
  
  if (empty ($rsa['kbmValue'])) {
    $TemplateCheckList->setValue('checkingRSA', 'Нулевой коэффициент');
  } else {
    $TemplateCheckList->setValue('checkingRSA', $rsa['kbmValue']);
  }
} elseif (array_key_exists('error', $rsa)){
    $er = 'Ошибка '.$rsa['error'].': '.$rsa['message'];
    $TemplateCheckList->setValue('checkingRSA', $er);
  } else {
    $TemplateCheckList->setValue('checkingRSA', 'Неизвестный формат ответа');
  };



//////ФЕДРЕСУРС



$fio = urlencode($surname.' '.$firstname.' '.$fathername);
curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api-cloud.ru/api/bankrot.php?type=searchString&string={$fio}&legalStatus=fiz&token=#############",//Токен можно получить тут https://api-cloud.ru/mvd
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'GET',
));

$fedresurs = json_decode(curl_exec($curl), true);

if(array_key_exists('status', $fedresurs)){
  $TemplateCheckList->setValue('checkingFedresurs', $fedresurs['message']);
} elseif (array_key_exists('error', $fedresurs)){
  $er = 'Ошибка '.$fedresurs['error'].': '.$fedresurs['message'];
  $TemplateCheckList->setValue('checkingFedresurs', $er);
} else {
  $TemplateCheckList->setValue('checkingFedresurs', 'Неизвестный формат ответа');
};



//////ФССП



curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api-cloud.ru/api/fssp.php?type=physical&lastname={$surname}&firstname={$firstname}&region={$homeRegion}&token=#############",//Токен можно получить тут https://api-cloud.ru/mvd
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'GET',
));

$fssp = json_decode(curl_exec($curl), true);

if(array_key_exists('status', $fssp)){
   $TemplateCheckList->setValue('ChekingFSSP', $fssp['message']);
  
} elseif (array_key_exists('error', $fssp)){
  $er = 'Ошибка '.$fssp['error'].': '.$fssp['message'];
  $TemplateCheckList->setValue('ChekingFSSP', $er);
} else {
  $TemplateCheckList->setValue('ChekingFSSP', 'Неизвестный формат ответа');
};



//////ИНКВИЗИТОР



$tgmMessege = urlencode("$surname $firstname $fathername $birthday $fullDDnum $issueDateDD $fullPasNum $phonenumber");
curl_setopt_array($curl, array(
  CURLOPT_URL => "https://api.telegram.org/bot5551608925:AAER3TG-73fJINzOpT-Gk_rlpprwMvDGVVU/sendMessage?chat_id=882504490&text={$tgmMessege}",
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'POST',
));

$tgm = json_decode(curl_exec($curl), true);
//print_r ($tgm['ok']);


curl_close($curl);

//Сохраняем чек-лист

$pathToSaveCheck = "C:\Users\ПК\Documents\\test\\$setData чек $surname.docx"; //Задает путь для сохранения файла
$TemplateCheckList->saveAs($pathToSaveCheck); //Сохраняет файл

//Заполняем и отправляем на почту заявку на скоринг

$TemplateOrderSK->setValue('fullname', $fullname);
$TemplateOrderSK->setValue('birthday', $birthday);
$TemplateOrderSK->setValue('PasNumSeria', $fullPasNum);
$TemplateOrderSK->setValue('issueDatePas', $issueDatePas);
$TemplateOrderSK->setValue('divCodePas', $pasportDivisionCode);
$TemplateOrderSK->setValue('DDNumSeria', $fullDDnum);
$TemplateOrderSK->setValue('issueDateDD', $issueDateDD);
$TemplateOrderSK->setValue('finalDateDD', $deadlineDateDD);
$TemplateOrderSK->setValue('drivingStage', $startDrivingYear);

$pathToSaveSKorder = "C:\Users\ПК\Documents\\test\Заявки в СК\\$surname.docx"; //Задает путь для сохранения файла
$TemplateOrderSK->saveAs($pathToSaveSKorder); //Сохраняет файл


//Отправка документа на почту

$mail->From = "IVashurin@yandex.ru";
$mail->FromName = "Вашурин Илья";

$mail->addAddress('IVashurin@yandex.ru', 'Илья Вашурин');
$mail->addAttachment($pathToSaveSKorder, 'Вашурин.docx');
$mail->Subject = 'Скоринг KIA Flex';
$mail->Body    = 'Александр, добрый день! Прошу провести скоринг водителей на допуск к выдаче автомобиля. Спасибо.';


$mail->send();


 

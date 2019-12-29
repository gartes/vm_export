<?php
	namespace Plg\Vm_export\Models;
	
	class Users{
		
		public $Group_id = 1 ;

//		public $offset = 2200 ;
//		public $LimitStr = 100 ;
		public $offset = 0 ;
		public $LimitStr = 0 ;
		
		/**
		 * Users constructor.
		 * @since 3.9
		 */
		public function __construct ()
		{
			// Подключаем класс для работы с excel
			require_once(JPATH_PLUGINS . '/vmextended/vm_export/Libraries/PHPExcel.php');
			// Подключаем класс для вывода данных в формате excel
			require_once(JPATH_PLUGINS . '/vmextended/vm_export/Libraries/PHPExcel/Writer/Excel5.php');
			
			
		}
		
		public function getUsers(){
			
			$db = \JFactory::getDbo();
			$query = $db->getQuery( true ) ;
			
			$a_select=[
				'a.id ',
				'a.id AS id_users ',
				'a.name',
				'a.username',
				'a.email',
				'a.registerDate',
			];
			$query->select( $a_select );
			$query->from('#__users AS a');
			
			$vu_select =[
				'vu.virtuemart_user_id',
				'vu.name',
				'vu.last_name',
				'vu.first_name',
				'vu.middle_name',
				'vu.phone_1',
				'vu.phone_2',
				'vu.address_1',
				'vu.address_2',
				'vu.city',
				'vu.virtuemart_state_id',
				'vu.virtuemart_country_id',
				'vu.created_on',
			];
			$query->select( $vu_select );
			$query->leftJoin('#__virtuemart_userinfos AS vu ON a.id = vu.virtuemart_user_id') ;
			
			
//			$query->select( ' o.* ' );
			$query->leftJoin('#__virtuemart_orders AS o  ON a.id = o.virtuemart_user_id') ;
			
			$gl_select =[
				'gl.glavpunkt_select_servise_type_on',
				'gl.cityText',
				'gl.city_id',
			];
			$query->select( $gl_select );
			$query->leftJoin('#__virtuemart_shipment_plg_glavpunkt AS gl  ON o.order_number = gl.order_number') ;
			
			$whereArr = [
//				$db->quoteName('first_name') . ' <> ' . $db->quote('Имя') ,
//				$db->quoteName('first_name') . ' <> ' . $db->quote('вова') ,
//				$db->quoteName('first_name') . ' <> ' . $db->quote('Олег Николайчук') ,
//				$db->quoteName('last_name') . ' <> ' . $db->quote('http://fgaerkhk.ru/') ,
//				$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук222') ,
//				$db->quoteName('phone_1') . ' <> ' . $db->quote('545454') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('baassgyl@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('ddf@ty.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('sad@asd') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('asd@asd') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('rsrif@gmail.com') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('jj@jj.ss') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('nn@nn.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('tresartt@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('zzz@sdsfg.rtye') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('test@ewrwerwrewe.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('dfasjdfsjg@eru.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('123123qwe@rer.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('rttasdbjg@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('325345123@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('12312414@aasdasd.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('rewrwerj@marewr.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('mail@ewrw.ru') ,
				$db->quoteName('a.username') . ' <> ' . $db->quote('33344222211111112311122331231333') ,
				$db->quoteName('a.username') . ' <> ' . $db->quote('33344222211111111126666233333') ,
			];
			$query->where($whereArr) ;
			
			$query->group(' a.id ');
			
			$db->setQuery($query , $this->offset , $this->LimitStr );
			$res = $db->loadObjectList();
			
//			$res = [];
			
//			echo'<pre>';print_r( $res );echo'</pre>'.__FILE__.' '.__LINE__;
//			die(__FILE__ .' '. __LINE__ );
			
			return $res ;
		}
		
		/**
		 * Список заказов
		 * @return mixed
		 * @since 3.9
		 */
		private function getOrders(){
			$db = \JFactory::getDbo();
			$query = $db->getQuery( true ) ;
			
			$o_select=[
				'o.* ',
				'o.virtuemart_order_id AS Order_id ',
				
//				'o.virtuemart_order_id ',
//				'o.virtuemart_user_id ',
//				'o.order_number ',
//				'o.customer_number ',
//				'o.order_total ',
//				'o.order_create_invoice_pass ',
			];
			$query->select(  $o_select );
			$query->from('#__virtuemart_orders AS o ');
			
			$vu_select=[
				'vu.*',
			];
			$query->select($vu_select);
			$query->leftJoin('#__virtuemart_userinfos AS vu ON o.virtuemart_user_id = vu.virtuemart_user_id') ;
			
			$a_select=[
				'a.*',
			];
			$query->select($a_select);
			$query->leftJoin('#__users AS a ON o.virtuemart_user_id = a.id') ;
			
			$gl_select=[
				'gl.*',
			];
			$query->select($gl_select);
			$query->leftJoin('#__virtuemart_shipment_plg_glavpunkt AS gl  ON o.order_number = gl.order_number') ;
			
			$whereArr = [
				/*$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук123123') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Дмитрий Федосеев') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Федосеев') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Oleg') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Ленин') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Ленинaaaaaa') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('фывфыв') ,
				*/
				/*$db->quoteName('a.email') . ' <> ' . $db->quote('sad.net@yandex.ua') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('2132@re.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('23423@rer.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('1231231212test@ru.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('22@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('test@ewrwerwrewe.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('zzz@sdsfg.rtye') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('dfasjdfsjg@eru.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('kndr@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('hhhh@yannnnnnllllldssex.ua') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('dondon@mail.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('1123123233@11111.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('ttewe@55123') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('wdwerwefsdfsdf@dsafsdgfsddfgdsf.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('2340-9345-094356-945-656@123123123123.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('32540946-9345693-45234095@12432142354536.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('124309359435956435-64536-0@219324923-9423-04.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('2345-02345435-3450-@1243129234587234.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('45-9459-7-0567-9567-5-96@23098409354096.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('dfgdfgdfg@fdgd11223fgd.fgh') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('ftps.kod@gmail.com') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('2350-9435-453@1242345345345.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('qwrewr345345345@4234324.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('asdasdasd@213123123126456456.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('32-0549364589076456@1244355.ru') ,
				$db->quoteName('a.email') . ' <> ' . $db->quote('sad.net@yandex.ru') ,*/
				
//				$db->quoteName('first_name') . ' <> ' . $db->quote('Олег Николайчук') ,
//				$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук') ,
				
//				order_status
				$db->quoteName('order_status') . ' <> ' . $db->quote('X') ,
				
//				$db->quoteName('o.virtuemart_order_id') . ' = ' . $db->quote('2230') ,
			];
			$query->where( $whereArr );
			
			$query->group(' o.virtuemart_order_id ') ;
			
			$db->setQuery($query , $this->offset , $this->LimitStr );
			
			$res = $db->loadObjectList();
			
//			echo'<pre>';print_r( $res );echo'</pre>'.__FILE__.' '.__LINE__;
//			die(__FILE__ .' '. __LINE__ );
			return $res;
		}
		
		private function getOrdersProduct(){
			$db = \JFactory::getDbo();
			$query = $db->getQuery( true ) ;
			$query->select( ' p.* '  );
			$query->from('#__virtuemart_order_items AS p ');
			
//			order_status
			$whereArr = [
				$db->quoteName('order_status') . ' <> ' . $db->quote('X') ,
			];
			$query->where($whereArr);
			
			$db->setQuery($query , 0 , $this->LimitStr );
			return $db->loadObjectList();
		}
		
		
		public function getExselList($res){
			$StopFirst_name = [
				'Олег Николайчук',
				'Василий Пупкич',
				'тест тест тест',
				'Testec test',
				'фывфывwwqqasasasasasa',
				'фывпрапрапрапрапрфыв',
				'Ленинaaaaaa',
				'Тест тестовый',
				'фывфывssdasd',
				'фывфыв',
				'Николайчук Олег Леонидович',
				'GGEWR ASD',
				'Test Test',
				'wdcvssxc павапм',
				'ejejfijd wififudj',
				'аоаовл клаомлм',
				'sjdjdjdjd dhdhdj',
			];
			
			
			// Создаем объект класса PHPExcel
			$xls = new \PHPExcel();
			// Устанавливаем индекс активного листа
			$xls->setActiveSheetIndex(0);
			// Получаем активный лист
			$sheet = $xls->getActiveSheet();
			// Подписываем лист
			$sheet->setTitle('Шаг1 - Юзеры');
			
			// Вставляем текст в ячейку A1
			$sheet->setCellValue("A1", 'Customer ID');
			$sheet->getStyle('A1')->getFill()->setFillType( \PHPExcel_Style_Fill::FILL_SOLID);
			$sheet->getStyle('A1')->getFill()->getStartColor()->setRGB('EEEEEE');
			// Выравнивание текста
			$sheet->getStyle('A1')->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			# Заголовки
			try
			{
				// Code that may throw an Exception or Error.
				// Вставляем текст в ячейку A2
				$sheet->setCellValue( "B1" , 'Group id' );
				$sheet->setCellValue( "C1" , 'Store id' );
				$sheet->setCellValue( "D1" , 'Address id' );
				$sheet->setCellValue( "E1" , 'Language id' );
				$sheet->setCellValue( "F1" , 'First name' );
				$sheet->setCellValue( "G1" , 'Last name' );
				$sheet->setCellValue( "H1" , 'Email' );
				$sheet->setCellValue( "I1" , 'Telephone' );
				$sheet->setCellValue( "J1" , 'Fax' );
				$sheet->setCellValue( "K1" , 'Custom field' );
				$sheet->setCellValue( "L1" , 'Password' );
				$sheet->setCellValue( "M1" , 'Salt' );
				$sheet->setCellValue( "N1" , 'Newsletter' );
				$sheet->setCellValue( "O1" , 'Safe' );
				$sheet->setCellValue( "P1" , 'Cart' );
				$sheet->setCellValue( "Q1" , 'Wish list' );
				$sheet->setCellValue( "R1" , 'IP' );
				$sheet->setCellValue( "S1" , 'Token' );
				$sheet->setCellValue( "T1" , 'Code' );
				$sheet->setCellValue( "U1" , 'Status' );
				$sheet->setCellValue( "V1" , 'Date added' );
				$sheet->setCellValue( "W1" , 'Deleted' );
			}
			catch( \Exception $e )
			{
				// Executed only in PHP 5, will not be reached in PHP 7
				echo 'Выброшено исключение: ' , $e->getMessage() , "\n";
			}
			catch( \Throwable $e )
			{
				// Executed only in PHP 7, will not match in PHP 5
				echo 'Выброшено исключение: ' , $e->getMessage() , "\n";
				echo '<pre>';
				print_r( $e );
				echo '</pre>' . __FILE__ . ' ' . __LINE__;
			}
			
			
			// Объединяем ячейки
			// $sheet->mergeCells('A1:H1');
			
			$i = 2 ;
			foreach( $res as $iUser => $item )
			{
				$first_name = $this->getName( $item );
				$Phone = $this->getPhone( $item );
				if( in_array( $first_name , $StopFirst_name) ) continue ;
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i , $item->id_users ); #Customer ID
					$sheet->setCellValueByColumnAndRow( 1 , $i , $this->Group_id ); #Group id
					$sheet->setCellValueByColumnAndRow( 2 , $i , 0 ); #Store id
					$sheet->setCellValueByColumnAndRow( 3 , $i , $iUser+1 ); #Address id
					$sheet->setCellValueByColumnAndRow( 4 , $i , 1 ); #Language id
					$sheet->setCellValueByColumnAndRow( 5 , $i , $first_name ); #First name
					$sheet->setCellValueByColumnAndRow( 6 , $i , '' ); #Last name
					$sheet->setCellValueByColumnAndRow( 7 , $i , $item->email ); #Email
					$sheet->setCellValueByColumnAndRow( 8 , $i , $Phone ); #Telephone
					$sheet->setCellValueByColumnAndRow( 9 , $i , '' ); #Fax
					$sheet->setCellValueByColumnAndRow( 10 , $i , '' ); #Custom field
					$sheet->setCellValueByColumnAndRow( 11 , $i , '' ); #Password
					$sheet->setCellValueByColumnAndRow( 12 , $i , '' ); #Salt
					$sheet->setCellValueByColumnAndRow( 13 , $i , 0 ); #Newsletter
					$sheet->setCellValueByColumnAndRow( 14 , $i , 0 ); #Safe
					$sheet->setCellValueByColumnAndRow( 15 , $i , '' ); #Cart
					$sheet->setCellValueByColumnAndRow( 16 , $i , '' ); #Wish list
					$sheet->setCellValueByColumnAndRow( 17 , $i , '' ); #IP
					$sheet->setCellValueByColumnAndRow( 18 , $i , '' ); #Token
					$sheet->setCellValueByColumnAndRow( 19 , $i , '' ); #Code
					$sheet->setCellValueByColumnAndRow( 20 , $i , '1' ); #Status
					$sheet->setCellValueByColumnAndRow( 21 , $i , $item->registerDate ); #Date added
					$sheet->setCellValueByColumnAndRow( 22 , $i , 0 ); #Deleted
					
				}
				catch( \PHPExcel_Exception $e )
				{
				}
				
				$i++;
//				echo'<pre>';print_r( $item->city );echo'</pre>'.__FILE__.' '.__LINE__;
			}#END FOREACH
			
			### Шаг2-Адреса Юзеров #################################
			$xls->getActiveSheet();
			$sheet = $xls->createSheet(1);
			// Устанавливаем индекс активного листа
			$xls->setActiveSheetIndex(1);
			// Получаем активный лист
			$sheet = $xls->getActiveSheet();
			// Подписываем лист
			$sheet->setTitle('Шаг2-Адреса Юзеров');
			# Заголовки
			$headArr = [
				'Address id',
				'Customer id',
				'First name',
				'Last name',
				'Company',
				'Address 1',
				'Address 2',
				'Postcode',
				'City',
				'Zone id',
				'Country id',
				'Deleted',
				];
			foreach( $headArr as $i =>  $item )
			{
				$sheet->setCellValueByColumnAndRow( $i , 1 , $item );
				
			}#END FOREACH
			
			$i = 2 ;
			foreach( $res as $iUser => $item )
			{
				$retCity = $this->getCity($item);
				$first_name = $this->getName( $item );
				
				if( in_array( $first_name , $StopFirst_name) ) continue ;
				
				
				
				
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i , $iUser + 1 );
					$sheet->setCellValueByColumnAndRow( 1 , $i , $item->id_users );
					$sheet->setCellValueByColumnAndRow( 2 , $i , $first_name );   #First name
					
					$sheet->setCellValueByColumnAndRow( 3 , $i , '' );                  #Last name
					
					$sheet->setCellValueByColumnAndRow( 4 , $i , '' );                  #Company
//					$sheet->setCellValueByColumnAndRow( 5 , $i , $item->address_type_name );#Address 1
					$sheet->setCellValueByColumnAndRow( 5 , $i , '' );#Address 1
					$sheet->setCellValueByColumnAndRow( 6 , $i , '' );#Address 2
//					$sheet->setCellValueByColumnAndRow( 7 , $i , $item->zip );#Postcode
					$sheet->setCellValueByColumnAndRow( 7 , $i , '' );#Postcode
					
					$sheet->setCellValueByColumnAndRow( 8 , $i , $retCity['city'] );#City
					$sheet->setCellValueByColumnAndRow( 9 , $i ,  $retCity['ZoneID'] );#Zone id
					
					
					$sheet->setCellValueByColumnAndRow( 10 , $i ,  '176' );#Country id
					$sheet->setCellValueByColumnAndRow( 11 , $i ,  '0' );#Deleted
				}
				catch( \PHPExcel_Exception $e )
				{
				}
				$i++;
			}#END FOREACH
			
			
			#### Шаг3 - Заказы Юзеров ############################################
			$orders = $this->getOrders();
			
			try
			{
				$xls->getActiveSheet();
				$sheet = $xls->createSheet(2);
				// Устанавливаем индекс активного листа
				$xls->setActiveSheetIndex(2);
				// Получаем активный лист
				$sheet = $xls->getActiveSheet();
				// Подписываем лист
				$sheet->setTitle('Шаг3-Заказы Юзеров');
			}
			catch( \PHPExcel_Exception $e )
			{
			}
			
			$headArr = [
				'Order id',
				'Invoice no',
				'Invoice prefix',
				'Store id',
				'Store name',
				'Store url',
				'Customer id',
				'Customer group id',
				'Firstname',
				'Lastname',
				'Email',
				'Telephone',
				'Fax',
				'Custom field',
				'Payment firstname',
				'Payment lastname',
				'Payment company',
				'Payment address 1',
				'Payment address 2',
				'Payment city',
				'Payment postcode',
				'Payment country',
				'Payment country id',
				'Payment zone',
				'Payment zone id',
				'Payment address format',
				'Payment custom field',
				'Payment method',
				'Payment code',
				'Shipping firstname',
				'Shipping lastname',
				'Shipping company',
				'Shipping address 1',
				'Shipping address 2',
				'Shipping city',
				'Shipping postcode',
				'Shipping country',
				'Shipping country id',
				'Shipping zone',
				'Shipping zone id',
				'Shipping address format',
				'Shipping custom field',
				'Shipping method',
				'Shipping code',
				'Comment',
				'Total',
				'Order status id',
				'Affiliate id',
				'Commission',
				'Marketing id',
				'Tracking',
				'Language id',
				'Currency id',
				'Currency code',
				'Currency value',
				'Ip',
				'Forwarded ip',
				'User agent',
				'Accept language',
				'Date added',
				'Date modified',
			];
			foreach( $headArr as $i =>  $item )
			{
				try
				{
					$sheet->setCellValueByColumnAndRow( $i , 1 , $item );
				}
				catch( \PHPExcel_Exception $e )
				{
				}
			}#END FOREACH
			
			$uri = \Joomla\CMS\Uri\Uri::getInstance( 'SERVER' );
			
			
			
			$i = 2 ;
			foreach( $orders as  $item )
			{
				
				$retCity = $this->getCity($item);
				$first_name = $this->getName( $item );
				$Phone = $this->getPhone( $item );
				
				if( in_array( $first_name , $StopFirst_name) ) continue ;
				
				if( $item->email == 'arkadijgulaev720@gmail.com' )
				{
					echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
					die(__FILE__ .' '. __LINE__ );
				}#END IF
				
				
				
				
				
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i ,  $item->Order_id );
					$sheet->setCellValueByColumnAndRow( 1 , $i ,  '0' );
					$sheet->setCellValueByColumnAndRow( 2 , $i ,  $item->order_create_invoice_pass ); # Invoice prefix
					$sheet->setCellValueByColumnAndRow( 3 , $i ,  '0' ); # Store id
					$sheet->setCellValueByColumnAndRow( 4 , $i ,  'Benks Shop' ); # Store id
					$sheet->setCellValueByColumnAndRow( 5 , $i ,  $uri::base() ); # Store id
					$sheet->setCellValueByColumnAndRow( 6 , $i ,  $item->virtuemart_user_id ); # Customer id
					$sheet->setCellValueByColumnAndRow( 7 , $i ,  $this->Group_id ); # Customer group id
					
					$sheet->setCellValueByColumnAndRow( 8 , $i ,   $first_name ); # Firstname
					
					$sheet->setCellValueByColumnAndRow( 9 , $i ,  '' ); # Lastname
					
					$sheet->setCellValueByColumnAndRow( 10 , $i ,  $item->email /*$item->name*/ ); # Email
					$sheet->setCellValueByColumnAndRow( 11 , $i ,  $Phone ); # Telephone
					$sheet->setCellValueByColumnAndRow( 12 , $i ,  '' ); # Fax
					
					$sheet->setCellValueByColumnAndRow( 14 , $i ,  $first_name );               # Payment firstname
					$sheet->setCellValueByColumnAndRow( 15 , $i ,  '' );                        # Payment lastname
					
					$sheet->setCellValueByColumnAndRow( 19 , $i ,  $retCity['city'] );          # Payment city
					
					$sheet->setCellValueByColumnAndRow( 20 , $i ,  '' );                        # Payment postcode
					$sheet->setCellValueByColumnAndRow( 21 , $i ,  'Российская Федерация' );    # Payment country
					$sheet->setCellValueByColumnAndRow( 22 , $i ,  '176' );                     # Payment country id
					$sheet->setCellValueByColumnAndRow( 23 , $i ,  $retCity['city'] );               # Payment zone
					
					
					$sheet->setCellValueByColumnAndRow( 24 , $i ,  $retCity['ZoneID'] );                   # Payment zone id
					
					$sheet->setCellValueByColumnAndRow( 25 , $i ,  '' );                        # Payment address format
					$sheet->setCellValueByColumnAndRow( 26 , $i ,  '[]' );                      # Payment custom field
					$sheet->setCellValueByColumnAndRow( 27 , $i ,  'Оплата при получении' );    # Payment method
					$sheet->setCellValueByColumnAndRow( 28 , $i ,  'cod' );                     # Payment code
					
					
					$sheet->setCellValueByColumnAndRow( 29 , $i ,  $first_name );          # Shipping firstname
					$sheet->setCellValueByColumnAndRow( 30 , $i ,  '');                    # Shipping lastname
					
					$sheet->setCellValueByColumnAndRow( 34 , $i ,  $retCity['city'] );                # Shipping city
					$sheet->setCellValueByColumnAndRow( 35 , $i ,  '' );                         # Shipping postcode
					$sheet->setCellValueByColumnAndRow( 36 , $i ,  'Российская Федерация' );     # Shipping country
					$sheet->setCellValueByColumnAndRow( 37 , $i ,  '176' );                      # Shipping country id
					
					$sheet->setCellValueByColumnAndRow( 38 , $i ,  $retCity['city']  );          # Shipping zone
					$sheet->setCellValueByColumnAndRow( 39 , $i ,  $retCity['ZoneID'] );         # Shipping zone id
					
					$sheet->setCellValueByColumnAndRow( 40 , $i ,  '' );                         # Shipping address format
					$sheet->setCellValueByColumnAndRow( 41 , $i ,  '[]' );                       # Shipping custom field
					$sheet->setCellValueByColumnAndRow( 42 , $i ,  'Главпункт - самовывоз' ); # Shipping method
					$sheet->setCellValueByColumnAndRow( 43 , $i ,  'glavpunkt.glavpunkt' ); # Shipping code
					$sheet->setCellValueByColumnAndRow( 44 , $i ,  '' ); # Comment
					$sheet->setCellValueByColumnAndRow( 45 , $i ,  $item->order_subtotal ); # Total
					$sheet->setCellValueByColumnAndRow( 46 , $i ,  '1' ); # Order status id
					$sheet->setCellValueByColumnAndRow( 47 , $i ,  '0' ); # Affiliate id
					$sheet->setCellValueByColumnAndRow( 48 , $i ,  '0.0000' ); # Commission
					$sheet->setCellValueByColumnAndRow( 49 , $i ,  '0' ); # Marketing id
					$sheet->setCellValueByColumnAndRow( 50 , $i ,  '' ); # Tracking
					$sheet->setCellValueByColumnAndRow( 51 , $i ,  '1' ); # Language id
					$sheet->setCellValueByColumnAndRow( 52 , $i ,  '1' ); # Currency id
					$sheet->setCellValueByColumnAndRow( 53 , $i ,  'RUB' ); # Currency code
					$sheet->setCellValueByColumnAndRow( 54 , $i ,  '1.00000000' ); # Currency value
					$sheet->setCellValueByColumnAndRow( 55 , $i ,  $item->ip_address ); # Ip
					$sheet->setCellValueByColumnAndRow( 56 , $i ,  '' ); # Forwarded ip
					$sheet->setCellValueByColumnAndRow( 57 , $i ,  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.88 Safari/537.36' ); # User agent
					$sheet->setCellValueByColumnAndRow( 58 , $i ,  'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7' ); # Accept language
					$sheet->setCellValueByColumnAndRow( 59 , $i ,  $item->created_on ); # Date added
					$sheet->setCellValueByColumnAndRow( 60 , $i ,  $item->modified_on ); # Date modified
				}
				catch( \PHPExcel_Exception $e )
				{
				}
				$i++ ;
			}
			
			### Шаг4-Товары в заказах #######################################
			try
			{
				$xls->getActiveSheet();
				$sheet = $xls->createSheet(3);
				// Устанавливаем индекс активного листа
				$xls->setActiveSheetIndex(3);
				// Получаем активный лист
				$sheet = $xls->getActiveSheet();
				// Подписываем лист
				$sheet->setTitle('Шаг4-Товары в заказах');
			}
			catch( \PHPExcel_Exception $e )
			{
			}
			
			$headArr = [
				'Order product id',
				'Order id',
				'Product id',
				'Name',
				'Model',
				'Quantity',
				'Price',
				'Total',
				'Tax',
				'Reward',
			];
			foreach( $headArr as $i =>  $item )
			{
				try
				{
					$sheet->setCellValueByColumnAndRow( $i , 1 , $item );
				}
				catch( \PHPExcel_Exception $e )
				{
				}
			}#END FOREACH
			
			$resOrdersProduct = $this->getOrdersProduct();
			
			$i = 2 ;
			foreach( $resOrdersProduct as  $item )
			{
				
//				echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
//				die(__FILE__ .' '. __LINE__ );
				
				if( empty( $item->order_item_name ) )
				{
					continue ;
				}#END IF
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i , $item->virtuemart_order_item_id ); # Order product id
					$sheet->setCellValueByColumnAndRow( 1 , $i , $item->virtuemart_order_id ); # Order id
					$sheet->setCellValueByColumnAndRow( 2 , $i , $item->virtuemart_product_id ); # Product id
					$sheet->setCellValueByColumnAndRow( 3 , $i , $item->order_item_name ); # Name
					$sheet->setCellValueByColumnAndRow( 4 , $i , $item->order_item_sku ); # Model
					$sheet->setCellValueByColumnAndRow( 5 , $i , $item->product_quantity ); # Quantity
					$sheet->setCellValueByColumnAndRow( 6 , $i , $item->product_item_price ); # Price
					$sheet->setCellValueByColumnAndRow( 7 , $i , $item->product_subtotal_with_tax ); # Total
					$sheet->setCellValueByColumnAndRow( 8 , $i , '0.0000' ); # Tax
					$sheet->setCellValueByColumnAndRow( 9 , $i , '0' ); # Reward
					$i++ ;
				}
				catch( \PHPExcel_Exception $e )
				{
				}
			}
			
			
			### Export-Order_totals ###########################################
			try
			{
				$xls->getActiveSheet();
				$sheet = $xls->createSheet(4);
				// Устанавливаем индекс активного листа
				$xls->setActiveSheetIndex(4);
				// Получаем активный лист
				$sheet = $xls->getActiveSheet();
				// Подписываем лист
				$sheet->setTitle('Export-Order_totals');
			}
			catch( \PHPExcel_Exception $e )
			{
			}
			
			$headArr = [ 'Order total id' , 'Order id' , 'Code' , 'Title' , 'Value' , 'Sort order' , ];
			foreach( $headArr as $i =>  $item )
			{
				try
				{
					$sheet->setCellValueByColumnAndRow( $i , 1 , $item );
				}
				catch( \PHPExcel_Exception $e )
				{
				}
			}#END FOREACH
			
			$fieldArr = [ 'sub_total' , 'shipping' /*, 'tax'*/ , 'discounts', 'total' , ];
			$str = 2 ;
			$Order_total_id = 1 ;
			
			
			
//			echo'<pre>';print_r( $orders[299] );echo'</pre>'.__FILE__.' '.__LINE__;
//			die(__FILE__ .' '. __LINE__ );

/*
			
			order_billDiscountAmount - order_discount
ABS(-230,70 - 200) = 30.70 ( Y )
X = order_total * 100% / Y = До целого числа = 1%
			
			
			*/
   
//			echo'<pre>';print_r( $orders );echo'</pre>'.__FILE__.' '.__LINE__;
//			die(__FILE__ .' '. __LINE__ );
			foreach( $orders as $i => $item )
			{
				$first_name = $this->getName( $item );
				if( in_array( $first_name , $StopFirst_name) ) continue ;
				
				# Order_id 3723
				if( (int)$item->Order_id == 8836  )
				{
					
					
					/*$Y =  $item->order_salesPrice - $item->order_total  ;
					if( $Y > 0 )
					{
						$X  = 100 - ($item->order_total / ( $item->order_salesPrice/100 )) ;
					}
					echo'<pre>';print_r( 'Y = '. $Y  );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'ceil(Y) = '.ceil($Y) );echo'</pre>'.__FILE__.' '.__LINE__;
					
					echo'<pre>';print_r( 'X = '. $X  );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'ceil(X)  = '. ceil($X) );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'round(X)  = '. round($X) );echo'</pre>'.__FILE__.' '.__LINE__;
					
					
					echo'<pre>';print_r( 'Order_id '.$item->Order_id );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_total '.$item->order_total );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_salesPrice '.$item->order_salesPrice );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_discountAmount '.$item->order_discountAmount );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_billDiscountAmount '.$item->order_billDiscountAmount );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_shipment '. $item->order_shipment );echo'</pre>'.__FILE__.' '.__LINE__;
					echo'<pre>';print_r( 'order_discount '. (int)$item->order_discount );echo'</pre>'.__FILE__.' '.__LINE__;
					
					
					echo'<pre>';print_r( 'order_total / 100 == '. $item->order_total / 100  );echo'</pre>'.__FILE__.' '.__LINE__;
					
					echo'<pre>';print_r( 'order_billDiscountAmount - order_discount == '. $item->order_discount );echo'</pre>'.__FILE__.' '.__LINE__;
					
					echo'<pre>';print_r( $orders[$i] );echo'</pre>'.__FILE__.' '.__LINE__;
					
					
					
					die(__FILE__ .' '. __LINE__ );*/
					
					
				}#END IF
//				echo'<pre>';print_r( $item->order_discount );echo'</pre>'.__FILE__.' '.__LINE__;
//				die(__FILE__ .' '. __LINE__ );
				
				foreach( $fieldArr as $id => $field )
				{
					try
					{
						switch($field){
							case 'sub_total':
								$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
								$sheet->setCellValueByColumnAndRow( 1 , $str , $item->Order_id ); # Order id
								$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Предварительная стоимость' ); # Title
//								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_total ); # Value
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_salesPrice ); # Value
								$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
								$str ++;
								$Order_total_id++;
								
								break ;
							case 'shipping':
								$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
								$sheet->setCellValueByColumnAndRow( 1 , $str , $item->Order_id ); # Order id
								$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Фиксированная стоимость доставки' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_shipment ); # Value
								$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
								$str ++;
								$Order_total_id++;
								
								break ;
							case 'tax':
								/*$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
								$sheet->setCellValueByColumnAndRow( 1 , $str , $item->Order_id ); # Order id
								$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'НДС (20%)' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_tax ); # Value
								$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
								$str ++;
								$Order_total_id++;*/
								
								break ;
							case 'total':
								$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
								$sheet->setCellValueByColumnAndRow( 1 , $str , $item->Order_id ); # Order id
								$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Итого' ); # Title
								# $sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_salesPrice ); # Value
								$sheet->setCellValueByColumnAndRow( 4 , $str , ceil ($item->order_total) ); # Value
								$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
								$str ++;
								$Order_total_id++;
								
								break ;
							case 'discounts':
								
								$Y =  $item->order_salesPrice - $item->order_total  ;
								if( $Y > 0 ){
									
									$X  = 100 - ($item->order_total / ( $item->order_salesPrice/100 )) ;
									// $X  =  ($item->order_salesPrice / $Y )/100    ;
									
									$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
									$sheet->setCellValueByColumnAndRow( 1 , $str , $item->Order_id ); # Order id
									$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
									$sheet->setCellValueByColumnAndRow( 3 , $str , 'Накопительная скидка и скидка от суммы заказа ('.round($X).'%)' ); # Title
									$sheet->setCellValueByColumnAndRow( 4 , $str , ceil($Y) ); # Value
									$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
									$str ++;
									$Order_total_id++;
									
									
//									die(__FILE__ .' '. __LINE__ );
									
								}
								
								break;
						}#END SWITCH
						
						
					}
					catch( \PHPExcel_Exception $e )
					{
					}
					
				}#END FOREACH
				
//				die(__FILE__ .' '. __LINE__ );
				
			}
			
			
			
			
			
//			echo'<pre>';print_r( $orders );echo'</pre>'.__FILE__.' '.__LINE__;
//			die(__FILE__ .' '. __LINE__ );
			
			
			// Выводим HTTP-заголовки
			header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
			header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
			header ( "Cache-Control: no-cache, must-revalidate" );
			header ( "Pragma: no-cache" );
			header ( "Content-type: application/vnd.ms-excel" );
			header ( "Content-Disposition: attachment; filename=matrix.xls" );
			
			// Выводим содержимое файла
			$objWriter = new \PHPExcel_Writer_Excel5($xls);
			$objWriter->save('php://output');
			
		}
		
		
		
		
		/**
		 * @param $item
		 *
		 * @return string
		 * @since 3.9
		 */
		private function getName ( $item )
		{
			if( !empty( $item->first_name ) && !empty( $item->last_name ) )
			{
				$first_name = $item->first_name . ' ' . $item->last_name;
			}
			else
			{
				$first_name = ( !empty( $item->first_name ) ? $item->first_name : $item->last_name );
			}
			
			return $first_name;#END IF
		}
		
		/**
		 * @param $item
		 *
		 * @return mixed
		 * @since 3.9
		 */
		private function getPhone ( $item )
		{
			return ( !empty( $item->phone_1 ) ? $item->phone_1 : $item->phone_2 );
		}
		
		/**
		 * @param $item
		 *
		 * @return array
		 * @since 3.9
		 */
		private function getCity ( $item )
		{
			$item->city = trim( $item->city );
			
			//				Москва
			//				Санкт-Петербург
			$retCity = [];
			$retCity['ZoneID'] = 2761;
			$retCity['city'] = 'Москва';
			if( stristr( $item->city , 'Санкт-Петербург' ) )
			{
				$retCity['ZoneID'] = 2785;
				$retCity['city'] = 'Санкт-Петербург' ;
			}
			else
			{
				if( empty( $item->city ) )
				{
					switch($item->glavpunkt_select_servise_type_on){
						case 'MSK':
							$retCity['city'] = 'Москва';
							$retCity['ZoneID'] = 2761;
							break ;
						
						case 'SPB':
							$retCity['ZoneID'] = 2785;
							$retCity['city'] = 'Санкт-Петербург' ;
							break ;
						default:
							$retCity['city'] = 'Москва';
							$retCity['ZoneID'] = 2761;
					}
				}#END IF
				
			}#END IF
			
			
			
			
			$stopArr = [
				'Комсомольск-на-Амуре',
				'Ростов-на-Дону',
				'Нижний Новгород',
			];
			$stopArrid_users = [
				'688',
				'689',
			];
			$stopArrFirst_name = [
				'Лесниченко Дмитрий Александрович',
			];
			/*if( !empty($item->glavpunkt_select_servise_type_on) && !in_array($item->city , $stopArr) && !in_array($item->id_users , $stopArrid_users) && !in_array($item->first_name , $stopArrFirst_name)  )
			{
//				echo'<pre>';print_r( $retCity );echo'</pre>'.__FILE__.' '.__LINE__;
//				echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
//				die(__FILE__ .' '. __LINE__ );
			}#END IF*/
			
			
			
			return $retCity;
		}
		
		
	}

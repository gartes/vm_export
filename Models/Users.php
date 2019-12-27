<?php
	namespace Plg\Vm_export\Models;
	
	class Users{
		
		public $Group_id = 1 ;
		public $LimitStr = 100 ;
//		public $LimitStr = 0 ;
		
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
			
			$query->select( 'a.* , vu.*' );
			$query->from('#__users AS a');
			$query->leftJoin('#__virtuemart_userinfos AS vu ON a.id = vu.virtuemart_user_id') ;
			
			$query->group(' a.id ');
			
			$db->setQuery($query);
			$res = $db->loadObjectList(); 
			
			$this->getExselList($res);

//			echo'<pre>';print_r( $res );echo'</pre>'.__FILE__.' '.__LINE__;
//		 	die(__FILE__ .' '. __LINE__ );
		}
		
		/**
		 * Список заказов
		 * @return mixed
		 * @since 3.9
		 */
		private function getOrders(){
			$db = \JFactory::getDbo();
			$query = $db->getQuery( true ) ;
			$query->select( ' o.* ,o.virtuemart_order_id AS Order_id , vu.* , a.*'  );
			$query->from('#__virtuemart_orders AS o ');
			$query->leftJoin('#__virtuemart_userinfos AS vu ON o.virtuemart_user_id = vu.virtuemart_user_id') ;
			$query->leftJoin('#__users AS a ON o.virtuemart_user_id = a.id') ;
			
			$whereArr = [
				$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Олег Николайчук123123') ,
				$db->quoteName('last_name') . ' <> ' . $db->quote('Дмитрий Федосеев') ,
			];
			$query->where( $whereArr );
			
			$query->group('o.virtuemart_order_id') ;
			
			$db->setQuery($query , 0 , $this->LimitStr );

			
			return $db->loadObjectList();
		}
		
		private function getOrdersProduct(){
			$db = \JFactory::getDbo();
			$query = $db->getQuery( true ) ;
			$query->select( ' p.* '  );
			$query->from('#__virtuemart_order_items AS p ');
			
			$db->setQuery($query , 0 , $this->LimitStr );
			return $db->loadObjectList();
		}
		
		
		private function getExselList($res){
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
			
			// Вставляем текст в ячейку A2
			$sheet->setCellValue("B1", 'Group id');
			$sheet->setCellValue("C1", 'Store id');
			$sheet->setCellValue("D1", 'Address id');
			$sheet->setCellValue("E1", 'Language id');
			$sheet->setCellValue("F1", 'First name');
			$sheet->setCellValue("G1", 'Last name');
			$sheet->setCellValue("H1", 'Email');
			$sheet->setCellValue("I1", 'Telephone');
			$sheet->setCellValue("J1", 'Fax');
			$sheet->setCellValue("K1", 'Custom field');
			$sheet->setCellValue("L1", 'Password');
			$sheet->setCellValue("M1", 'Salt');
			$sheet->setCellValue("N1", 'Newsletter');
			$sheet->setCellValue("O1", 'Safe');
			$sheet->setCellValue("P1", 'Cart');
			$sheet->setCellValue("Q1", 'Wish list');
			$sheet->setCellValue("R1", 'IP');
			$sheet->setCellValue("S1", 'Token');
			$sheet->setCellValue("T1", 'Code');
			$sheet->setCellValue("U1", 'Status');
			$sheet->setCellValue("V1", 'Date added');
			$sheet->setCellValue("W1", 'Deleted');
			
			
			
			// Объединяем ячейки
			// $sheet->mergeCells('A1:H1');
			
			$i = 2 ;
			foreach( $res as $iUser => $item )
			{
				$first_name = $this->getName( $item );
				
				$Phone = $this->getPhone( $item );
				
				if( $item->id == 704 )
				{
//					echo'<pre>';print_r($item    );echo'</pre>'.__FILE__.' '.__LINE__;
//					die(__FILE__ .' '. __LINE__ );
				}#END IF
				
				
				
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i , $item->id ); #Customer ID
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
//				Москва
//				Санкт-Петербург
				$ZoneID = 2761 ;
				if( stristr($item->city, 'Санкт-Петербург') )
				{
					$ZoneID = 2785 ;
				}else{
					$item->city = 'Москва' ;
				}#END IF
				
				$first_name = $this->getName( $item );
				
				if( $item->id == 680 )
				{
//					echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
//					die(__FILE__ .' '. __LINE__ );
				}#END IF
				
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i , $iUser + 1 );
					$sheet->setCellValueByColumnAndRow( 1 , $i , $item->id );
					$sheet->setCellValueByColumnAndRow( 2 , $i , $first_name );   #First name
					
					$sheet->setCellValueByColumnAndRow( 3 , $i , '' );                  #Last name
					
					$sheet->setCellValueByColumnAndRow( 4 , $i , '' );                  #Company
//					$sheet->setCellValueByColumnAndRow( 5 , $i , $item->address_type_name );#Address 1
					$sheet->setCellValueByColumnAndRow( 5 , $i , '' );#Address 1
					$sheet->setCellValueByColumnAndRow( 6 , $i , '' );#Address 2
//					$sheet->setCellValueByColumnAndRow( 7 , $i , $item->zip );#Postcode
					$sheet->setCellValueByColumnAndRow( 7 , $i , '' );#Postcode
					$sheet->setCellValueByColumnAndRow( 8 , $i , $item->city );#City
					$sheet->setCellValueByColumnAndRow( 9 , $i ,  $ZoneID );#Zone id
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
				
				$ZoneID = 2761 ;
				if( stristr($item->city, 'Санкт-Петербург') )
				{
					$ZoneID = 2785 ;
				}else{
					$item->city = 'Москва' ;
				}#END IF
				
				if( $item->virtuemart_order_id == 80 )
				{
//					echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
//					die(__FILE__ .' '. __LINE__ );
				}#END IF
				
				$first_name = $this->getName( $item );
				$Phone = $this->getPhone( $item );
				
				if( $first_name == 'Дмитрий Федосеев' )
				{
					
					echo'<pre>';print_r( $item );echo'</pre>'.__FILE__.' '.__LINE__;
					die(__FILE__ .' '. __LINE__ );
					
					continue ;
				}#END IF
				
				
				
				
				try
				{
					$sheet->setCellValueByColumnAndRow( 0 , $i ,  $item->virtuemart_order_id );
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
					
					$sheet->setCellValueByColumnAndRow( 19 , $i ,  $item->city ); # Fax
					
					$sheet->setCellValueByColumnAndRow( 20 , $i ,  '' );                        # Payment postcode
					$sheet->setCellValueByColumnAndRow( 21 , $i ,  'Российская Федерация' );    # Payment country
					$sheet->setCellValueByColumnAndRow( 22 , $i ,  '176' );                     # Payment country id
					$sheet->setCellValueByColumnAndRow( 23 , $i ,  $item->city );               # Payment zone
					
					
					$sheet->setCellValueByColumnAndRow( 24 , $i ,  $ZoneID );                   # Payment zone id
					
					$sheet->setCellValueByColumnAndRow( 25 , $i ,  '' );                        # Payment address format
					$sheet->setCellValueByColumnAndRow( 26 , $i ,  '[]' );                      # Payment custom field
					$sheet->setCellValueByColumnAndRow( 27 , $i ,  'Оплата при получении' );    # Payment method
					$sheet->setCellValueByColumnAndRow( 28 , $i ,  'cod' );                     # Payment code
					
					
					$sheet->setCellValueByColumnAndRow( 29 , $i ,  $first_name );          # Shipping firstname
					$sheet->setCellValueByColumnAndRow( 30 , $i ,  '');                    # Shipping lastname
					
					$sheet->setCellValueByColumnAndRow( 34 , $i ,  $item->city );                # Shipping city
					$sheet->setCellValueByColumnAndRow( 35 , $i ,  '' );                         # Shipping postcode
					$sheet->setCellValueByColumnAndRow( 36 , $i ,  'Российская Федерация' );     # Shipping country
					$sheet->setCellValueByColumnAndRow( 37 , $i ,  '176' );                      # Shipping country id
					$sheet->setCellValueByColumnAndRow( 38 , $i ,  $item->city );                # Shipping zone
					$sheet->setCellValueByColumnAndRow( 39 , $i ,  $ZoneID );                    # Shipping zone id
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
			
			$fieldArr = [ 'sub_total' , 'shipping' , 'tax' , 'total' , ];
			$str = 2 ;
			$Order_total_id = 1 ;
			foreach( $orders as $i => $item )
			{
				
				foreach( $fieldArr as $id => $field )
				{
					try
					{
						$sheet->setCellValueByColumnAndRow( 0 , $str , $Order_total_id ); # Order total id
						$sheet->setCellValueByColumnAndRow( 1 , $str , $item->virtuemart_order_id ); # Order id
						$sheet->setCellValueByColumnAndRow( 2 , $str , $field ); # Code
						
						switch($field){
							case 'sub_total':
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Предварительная стоимость' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_total ); # Value
								break ;
							case 'shipping':
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Фиксированная стоимость доставки' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_shipment ); # Value
								break ;
							case 'tax':
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'НДС (20%)' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_tax ); # Value
								break ;
							case 'total':
								$sheet->setCellValueByColumnAndRow( 3 , $str , 'Итого' ); # Title
								$sheet->setCellValueByColumnAndRow( 4 , $str , $item->order_total ); # Value
								break ;
						}
						$sheet->setCellValueByColumnAndRow( 5 , $str , $id ); # Title
						
					}
					catch( \PHPExcel_Exception $e )
					{
					}
					$str ++;
					$Order_total_id++;
				}#END FOREACH
				
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
		
		
	}

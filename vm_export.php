<?php
/**
 * @package    vm_export
 *
 * @author     oleg <your@email.com>
 * @copyright  A copyright
 * @license    GNU General Public License version 2 or later; see LICENSE.txt
 * @link       http://your.url.com
 */

use Joomla\CMS\Application\CMSApplication;
use Joomla\CMS\Plugin\CMSPlugin;
use Joomla\Database\DatabaseDriver;

defined('_JEXEC') or die;

/**
 * Vm_export plugin.
 *
 * @package   vm_export
 * @since     1.0.0
 */
class plgVmExtendedVm_export extends vmExtendedPlugin
{
	/**
	 * Application object
	 *
	 * @var    CMSApplication
	 * @since  1.0.0
	 */
	protected $app;

	/**
	 * Database object
	 *
	 * @var    DatabaseDriver
	 * @since  1.0.0
	 */
	protected $db;

	/**
	 * Affects constructor behavior. If true, language files will be loaded automatically.
	 *
	 * @var    boolean
	 * @since  1.0.0
	 */
	protected $autoloadLanguage = true;
	
	/**
	 * plgVmExtendedVm_export constructor.
	 *
	 * @param          $subject
	 * @param   array  $config
	 * @since 3.9
	 */
	public function __construct (&$subject, $config=array()) {
		JLoader::registerNamespace(
			'Plg\Vm_export\Models' ,
			JPATH_PLUGINS . '/vmextended/vm_export/Models' ,
			$reset = false , $prepend = false , $type = 'psr4' );
		
		parent::__construct($subject, $config);
	}  // end function
	
	/**
	 * Точка входа для BE Admin
	 * $controller - index.php?option=com_virtuemart&view={$controller}
	 * @param   string  $controller
	 *
	 * @return True|void
	 * @since 3.9
	 */
	public function onVmAdminController ($controller){
	
	}
	
	/**
	 * Запуск эспорта покупателей для Open Cart
	 * index.php?option=com_virtuemart&view=vm_export
	 *
	 * @param   string  $controller
	 *
	 * @return True|void
	 * @throws Exception
	 * @since 3.9
	 */
	public function onVmSiteController ($controller){
		
		if( $controller != 'vm_export' ) {
			throw new Exception('Включен плагин VM_EXPORT - Запуск эспорта покупателей для Open Cart' , 1000500);
		}  #END IF
		
		$Users = new \Plg\Vm_export\Models\Users();
		$res = $Users->getUsers();
		$Users->getExselList($res);
	}
	
	
	/**
	 * onAfterInitialise.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterInitialise()
	{
		
	}

	/**
	 * onAfterRoute.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterRoute()
	{
	
	}

	/**
	 * onAfterDispatch.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterDispatch()
	{
	
	}

	/**
	 * onAfterRender.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterRender()
	{
		// Access to plugin parameters
		$sample = $this->params->get('sample', '42');
	}

	/**
	 * onAfterCompileHead.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterCompileHead()
	{
	
	}

	/**
	 * OnAfterCompress.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterCompress()
	{
	
	}

	/**
	 * onAfterRespond.
	 *
	 * @return  void
	 *
	 * @since   1.0.0
	 */
	public function onAfterRespond()
	{
	
	}
}

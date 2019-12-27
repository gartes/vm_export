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
class plgSystemVm_export extends CMSPlugin
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

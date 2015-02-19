<?php
/**
 * CakePHP(tm) : Rapid Development Framework (http://cakephp.org)
 * Copyright 2005-2012, Cake Software Foundation, Inc. (http://cakefoundation.org)
 *
 * Licensed under The MIT License
 * Redistributions of files must retain the above copyright notice.
 *
 * @copyright     Copyright 2005-2012, Cake Software Foundation, Inc. (http://cakefoundation.org)
 * @link          http://cakephp.org CakePHP(tm) Project
 * @license       MIT License (http://www.opensource.org/licenses/mit-license.php)
 */

App::uses('View', 'View');

/**
 * @package       CakePHPExcel.View
 */
class ExcelView extends View {

	/**
	 * The subdirectory. XLS views are always in xls.
	 *
	 * @var string
	 */
	public $subDir = 'xls';

	/**
	 * Default excel configs.
	 *
	 * @var string
	 */
	public $excelConfig = array(
		'creator' => 'CakePHPExcel <http://github.com/maraya/CakePHPExcel>'
	);

	/**
	 * Excel format
	 *
	 * @var string
	 */
	public $format;

	/**
	 * Excel extension
	 *
	 * @var string
	 */
	public $ext;

	/**
	 * Constructor
	 *
	 * @param Controller $controller
	 * @return void
	 */
	public function __construct(Controller $controller = null) {
		$this->excelConfig = array_merge(
			(array)$this->excelConfig,
			(array)$controller->excelConfig
		);

		$this->ext = $controller->request->params['ext'];
		$this->format = ($this->ext == 'xlsx')? 'Excel2007': 'Excel5';

		if (!class_exists('PHPExcel_IOFactory')) {
			App::import('Vendor', 'IOFactory', array('file' => 'phpoffice' . DS . 'phpexcel' . DS . 'Classes' . DS . 'PHPExcel' . DS . 'IOFactory.php'));
		}

		parent::__construct($controller);
	}

	/**
	 * Render an Excel for download.
	 *
	 * @param string $view The view being rendered.
	 * @param string $layout The layout being rendered.
	 * @return string The rendered view.
	 */
	public function render($view = null, $layout = null) {
		$this->layoutPath = 'xls';
		$content = parent::render($view, $layout);

		$file = $this->createTempFile($content);
		$reader = PHPExcel_IOFactory::createReader('HTML');

		$excel = $reader->load($file);
		$this->deleteTempFile($file);

		$props = $excel->getProperties();
		$props->setCreator($this->excelConfig['creator']);
		
		$writer = PHPExcel_IOFactory::createWriter($excel, $this->format);
		ob_start();
		$writer->save('php://output');
		$excelOutput = ob_get_clean();

		if (isset($this->excelConfig['filename'])) {
			$this->response->download($this->excelConfig['filename']);
		}

		$this->Blocks->set('content', $excelOutput);
		return $this->Blocks->get('content');
	}

	/**
	 * Create temp file in sys_get_temp_dir with view content.
	 *
	 * @param string $content The view content
	 * @return string The tempfile path.
	 */
	private function createTempFile($content) {
		$file = tempnam(sys_get_temp_dir(), 'cakephpexcel_');
    	$fp = fopen($file, "w");
    	fwrite($fp, $content);
    	fclose($fp);
    	return $file;
	}

	/**
	 * Delete tempfile
	 *
	 * @param string $file The tempfile path
	 * @return void
	 */
	private function deleteTempFile($file) {
		unlink($file);
	}
}
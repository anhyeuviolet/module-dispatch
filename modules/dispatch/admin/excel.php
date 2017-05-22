<?php

/**
 * @Project NUKEVIET 4.x
 * @Author VINADES (contact@vinades.vn)
 * @Copyright (C) 2014 VINADES. All rights reserved
 * @License GNU/GPL version 2 or any later version
 * @Createdate 2-9-2010 14:43
 */
if (! defined('NV_IS_FILE_ADMIN')) {
    die('Stop!!!');
}
$from = $to = $cat = $type = 0;

$listcats = nv_listcats(0);

if (empty($listcats)) {
    Header("Location: " . NV_BASE_ADMINURL . "index.php?" . NV_LANG_VARIABLE . "=" . NV_LANG_DATA . "&" . NV_NAME_VARIABLE . "=" . $module_name . "&" . NV_OP_VARIABLE . "=cat&add=1");
    exit();
}

$listtypes = nv_listtypes($type, 0);

if (empty($listtypes)) {
    Header("Location: " . NV_BASE_ADMINURL . "index.php?" . NV_LANG_VARIABLE . "=" . NV_LANG_DATA . "&" . NV_NAME_VARIABLE . "=" . $module_name . "&" . NV_OP_VARIABLE . "=tyes&add=1");
    exit();
}

$xtpl = new XTemplate("excel.tpl", NV_ROOTDIR . "/themes/" . $global_config['module_theme'] . "/modules/" . $module_file);
$xtpl->assign('LANG', $lang_module);
$xtpl->assign('NV_BASE_ADMINURL', NV_BASE_ADMINURL);
$xtpl->assign('NV_BASE_SITEURL', NV_BASE_SITEURL);
$xtpl->assign('NV_LANG_INTERFACE', NV_LANG_INTERFACE);
$xtpl->assign('NV_NAME_VARIABLE', NV_NAME_VARIABLE);
$xtpl->assign('NV_OP_VARIABLE', NV_OP_VARIABLE);
$xtpl->assign('NV_ASSETS_DIR', NV_ASSETS_DIR);
$xtpl->assign('MODULE_NAME', $module_name);
$xtpl->assign('OP', $op);
$xtpl->assign('FILES_DIR', NV_UPLOADS_DIR . '/' . $module_upload);
$xtpl->assign('FORM_ACTION', NV_BASE_ADMINURL . "index.php?" . NV_LANG_VARIABLE . "=" . NV_LANG_DATA . "&amp;" . NV_NAME_VARIABLE . "=" . $module_name . "&amp;" . NV_OP_VARIABLE . "=" . $op);

foreach ($listtypes as $types) {
    $xtpl->assign('LISTTYPES', $types);
    $xtpl->parse('main.typeid');
}

foreach ($listcats as $cats) {
    $xtpl->assign('LISTCATS', $cats);
    $xtpl->parse('main.catid');
}

if( $nv_Request->isset_request("export", "get") ){
	if (!defined('NV_IS_AJAX')) die('Wrong URL');
	

	$where_export = '';
	// By Type
	if ($nv_Request->isset_request("type", "get")) {
		$type = $nv_Request->get_int('type', 'get', 0);
		if($type > 0){
			$where_export .= " AND type=" . $type;
		}
	}

	// By Cat
	if ($nv_Request->isset_request("catid", "get")) {
		$catid = $nv_Request->get_int('catid', 'get', 0);
		if($catid > 0){
			$where_export .= " AND catid=" . $catid;
		}
	}

	$db->sqlreset()
		->select('COUNT(*)')
		->from(NV_PREFIXLANG . '_' . $module_data . '_document WHERE id!=0' . $where_export);

	$num_items = $db->query($db->sql())
		->fetchColumn();
	$listtypes = nv_listtypes(0);

		$db->select('*')
			->order('type ASC, from_time ASC');
		$result = $db->query($db->sql())->fetchAll();

		$i = 2;
		//Goi thu vien PHPExcel
		// require_once NV_ROOTDIR . '/modules/' . $module_file . '/PHPExcel/PHPExcel.php';
		$objPHPExcel = new PHPExcel();
		// Khoi tao header cua cac cot chua du lien
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A1', $lang_module['cat'])
			->setCellValue('B1', $lang_module['dis'])
			->setCellValue('C1', $lang_module['dis_code'])
			->setCellValue('D1', $lang_module['dis_date_re'])
			->setCellValue('E1', $lang_module['date_delivery'])
			->setCellValue('F1', $lang_module['dis_souce'])
			->setCellValue('G1', $lang_module['copy_count'])
			->setCellValue('H1', $lang_module['dis_person'])
			->setCellValue('I1', $lang_module['dis_to_org'])
			->setCellValue('J1', $lang_module['dis_de'])
			->setCellValue('K1', $lang_module['dis_date_iss'])
			->setCellValue('L1', $lang_module['dis_date_first'])
			->setCellValue('M1', $lang_module['dis_date_die'])
			->setCellValue('N1', $lang_module['dis_content']);
		foreach($result as $row){
			//set gia tri cho cac cot du lieu
			
			$row['signer'] = '';
			$sql = "SELECT * FROM " . NV_PREFIXLANG . "_" . $module_data . "_signer WHERE id=" . $row['from_signer'];
			$re = $db->query($sql);
			if ($re->rowCount()) {
				while ($ro = $re->fetch()) {
					$row['signer'] = $ro['name'];
				}
			}
			
			$row['dis_de'] = array();
			$sql = "SELECT * FROM " . NV_PREFIXLANG . "_" . $module_data . "_de_do WHERE doid=" . $row['id'];
			$re = $db->query($sql);
			if ($re->rowCount()) {
				while ($ro = $re->fetch()) {
					$listdes = nv_listdes($ro['deid']);
					$row['dis_de'][] = $listdes[$ro['deid']]['title'];
				}
			}
			
			$row['from_time'] = nv_date("d/m/Y", $row['from_time']);
			$from_time = PHPExcel_Style_NumberFormat::toFormattedString($row['from_time'], "D/M/YYYY");
			
			$row['date_delivery'] = nv_date("d/m/Y", $row['date_delivery']);
			$date_delivery = PHPExcel_Style_NumberFormat::toFormattedString($row['date_delivery'], "D/M/YYYY");
			
			$row['date_iss'] = nv_date("d/m/Y", $row['date_iss']);
			$date_iss = PHPExcel_Style_NumberFormat::toFormattedString($row['date_iss'], "D/M/YYYY");
			
			$row['date_first'] = nv_date("d/m/Y", $row['date_first']);
			$date_first = PHPExcel_Style_NumberFormat::toFormattedString($row['date_first'], "D/M/YYYY");
			
			$row['date_die'] = nv_date("d/m/Y", $row['date_die']);
			$date_die = PHPExcel_Style_NumberFormat::toFormattedString($row['date_die'], "D/M/YYYY");
			
			$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A'.$i, $listcats[$row['catid']]['title'])
			->setCellValue('B'.$i, $listtypes[$row['type']]['title'])
			->setCellValue('C'.$i, $row['code'])
			->setCellValue('D'.$i, $from_time)
			->setCellValue('E'.$i, $date_delivery)
			->setCellValue('F'.$i, $row['from_org'])
			->setCellValue('G'.$i, $row['copy_count'])
			->setCellValue('H'.$i, $row['signer'])
			->setCellValue('I'.$i, $row['to_org'])
			->setCellValue('J'.$i, implode(',', $row['dis_de']))
			->setCellValue('K'.$i, $date_iss)
			->setCellValue('L'.$i, $date_first)
			->setCellValue('M'.$i, $date_die)
			->setCellValue('N'.$i, $row['content']);
			$i++;
		}

		$_name = $module_name . '_' . date('H_i_d_m_Y') . '.xls'; // ten va dinh dang file
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5'); // ghi du lieu vao file,dinh dang file Excel 2007 (Excel2007), CSV hoac Excel 2003(Excel5)
		
		$styleArray = array(
		  'borders' => array(
			'allborders' => array(
			  'style' => PHPExcel_Style_Border::BORDER_THIN
			)
		  )
		);

		$objPHPExcel->getActiveSheet()->getStyle('A1:N' . ($i - 1))->applyFromArray($styleArray);

		foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {

			$objPHPExcel->setActiveSheetIndex($objPHPExcel->getIndex($worksheet));

			$sheet = $objPHPExcel->getActiveSheet();
			$cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(true);
			/** @var PHPExcel_Cell $cell */
			foreach ($cellIterator as $cell) {
				$sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
			}
		}
		
		$objPHPExcel->getActiveSheet()->getStyle('A1:N1')->getFont()->setBold(true);
		unset($styleArray);

		$full_path = NV_ROOTDIR .'/'. NV_UPLOADS_DIR .'/'. $module_upload . '/' . $_name;//duong dan tuyet doi cua file
		$objWriter->save($full_path);
		// Dieu huong den duong link ben ngoai site de co the download
		Header('Location: ' . nv_url_rewrite( NV_MY_DOMAIN . NV_BASE_SITEURL . NV_UPLOADS_DIR .'/' . $module_upload . '/' . $_name ));
}
$page_title = $lang_module['excel'];

$xtpl->parse('main');
$contents = $xtpl->text('main');

include NV_ROOTDIR . '/includes/header.php';
echo nv_admin_theme($contents);
include NV_ROOTDIR . '/includes/footer.php';
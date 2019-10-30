<?php
/**
 * @package Cooling_importer
 * @version 1.0.0
 */
/*
Plugin Name: Cooling Importer
Plugin URI:
Description: Import Excel Data
Author:
Version:
Author URI:
*/
ini_set('display_errors', 'Off');

defined( 'ABSPATH' ) or die( 'No script kiddies please!' );

include( plugin_dir_path( __FILE__ ) . 'PHPExcel-1.8/Classes/PHPExcel.php');

add_action( 'admin_init', 'register_my_setting' );

function register_my_setting() {
  register_setting( 'excel-settings-group', 'data', 'update_excel' );
	register_setting('file-excel-group', 'urlexcel', 'upload_file');
}

if ( ! function_exists( 'wp_handle_upload' ) ) {
    require_once( ABSPATH . 'wp-admin/includes/file.php' );
}

function upload_file() {
//    add_filter( 'upload_dir', 'change_upload_dir' );
    $uploadedfile = $_POST['file_excel'];


  //  remove_filter( 'upload_dir', 'change_upload_dir' );
    return $uploadedfile;
}

function change_upload_dir( $dirs ) {
  return $dirs;
}

add_action( 'admin_menu', 'admin_excel' );

function admin_excel() {
	add_options_page( 'Excel importer', 'Excel importer', 'manage_options', 'menu_excel', 'create_excel_page' );
	// add_action( 'admin_init', 'register_settings_excel' );
}

function create_excel_page() {
	if ( !current_user_can( 'manage_options' ) )  {
		wp_die( __( 'You do not have sufficient permissions to access this page.' ) );
}
?>
<form enctype="multipart/form-data" method="post" action="options.php">
	<div style="font-size:24px; font-weight:bold;">
	<h3> 1. Upload Excel File </h3>
  </div>  <?php $url=get_option('urlexcel');
  ?>

	<div style="padding-bottom:20px"> <i>Last File Uploaded: <b> <?php echo $url; ?> </b></i></div>
	<div style="padding-bottom:20px">

  <input type="text" name="file_excel"value="<?php echo $url; ?>"  size="100" />

	</div>
	      <div> <input name="vai" type="submit" value="upload" class="button-primary"/> </div>
         <input type="hidden" name="action" value="upload_file" />

  <?php  settings_fields( 'file-excel-group' ); ?>
    <?php // do_settings( 'concorso-settings-group' ); ?>
  <?php // do_settings_sections(__FILE__); ?>
  </form>
	<form method="post" action="options.php">
	<div style="font-size:24px; font-weight:bold; padding-bottom:20px">
	<h3> 2. Update the database </h3>
    <div style="padding-bottom:20px">
       <span style='font-size:14px'>Import the first record </span><br/>
      <input type="checkbox" name="trial" />

    </div>
   <div> <input name="vai" type="submit" value="update" class="button-primary"/> </div>
	</div>
  <input type="hidden" name="action" value="update_excel" />

  <?php  settings_fields( 'excel-settings-group' ); ?>
  <?php // do_settings( 'concorso-settings-group' ); ?>  <?php // do_settings_sections(__FILE__); ?>
  <?php $data=get_option('data');

	?> <div> Last Update: <b> <?php echo date("d-m-Y G:i", $data); ?></b></div>
</form>
<?php
 }


 function filter_empty($val) {
  if(empty($val) || $val==" ") {
    unset($val);
  } else {
    return $val;
  }
}

 function update_excel () {
   $fileexcel=get_option('urlexcel');
  /*   $nomefile=explode("uploads/", $fileexcel);
    $perc=WP_CONTENT_DIR ."/uploads";
    $urlfinale=$perc."/".$nomefile[1]; */

    global $wpdb;

    define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

    file_put_contents('excel.xlsx',file_get_contents($fileexcel));

    $objReader = new PHPExcel_Reader_Excel2007();
    $objPHPExcel = $objReader->load("excel.xlsx");
    $x=$objPHPExcel->getActiveSheet()->getCell('C4')->getValue();
    $objWorksheet = $objPHPExcel->getActiveSheet();
    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);

    $trial= $_POST["trial"];

    if($trial=="on") {
      $lastRow=2;
    } else {
      $lastRow=$highestRow;
    }

    for ($row = 2; $row <= $lastRow; ++$row) {

          $sector=$objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
          $subsector=$objWorksheet->getCellByColumnAndRow(1, $row)->getValue();
          $titolo=addslashes($objWorksheet->getCellByColumnAndRow(2, $row)->getValue());
          $application=$objWorksheet->getCellByColumnAndRow(3, $row)->getValue();
          $tech=$objWorksheet->getCellByColumnAndRow(4, $row)->getValue();
          $refrigerant=$objWorksheet->getCellByColumnAndRow(5, $row)->getValue();
          $manufacturer=$objWorksheet->getCellByColumnAndRow(6, $row)->getValue();
          $country=$objWorksheet->getCellByColumnAndRow(7, $row)->getValue();
          $description=$objWorksheet->getCellByColumnAndRow(8, $row)->getValue();
          $energy=$objWorksheet->getCellByColumnAndRow(9, $row)->getValue();
          $post_type=strtolower($objWorksheet->getCellByColumnAndRow(10, $row)->getValue());
          if($post_type=="case study") {
            $post_type="case-study";
          }
          $web=$objWorksheet->getCellByColumnAndRow(11, $row)->getValue();
          $source=$objWorksheet->getCellByColumnAndRow(12, $row)->getValue();
          $tags=$objWorksheet->getCellByColumnAndRow(13, $row)->getValue();

          // FIND A POST WITH SOME TITLE
          $id=post_exists($titolo,'','',$post_type);

          $args=array("post_title"=>$titolo,'post_content'=>$description,'post_type'=>$post_type);

          // CREATE THE POST
          if(!$id) {
              $id=wp_insert_post($args);
          } else {
              $args["ID"]=$id;
            //  print_r($args);
              wp_update_post($args);
          }

          /* INSERT CATEGORIES */
          $sectors= explode("\n", $sector);
          $sectors=array_filter($sectors, "filter_empty");
          $numerisector=count($sectors);

          foreach($sectors as $s) {
            $s=ucfirst($s);
            $term_id = term_exists( $s, "type", $parent );
            if($term_id) {
              wp_set_post_terms( $id, $term_id, "type",true );
            }
          }
          $subsectors= explode("\n", $subsector);
          $subsectors= array_filter($subsectors,"filter_empty");

          foreach($subsectors as $su) {
            $su=ucfirst(strtolower($su));
            if($su=="Domestic-a") { $su="domestic-air-conditioning";}
            if($su=="Transport-a") { $su="transport-air-conditioning";}
            if($su=="Industrial-a") { $su="industrial-air-conditioning";}
            if($su=="Commercial-a") { $su="commercial-air-conditioning";}
            if($su=="Domestic-r") { $su="domestic-refrigeration";}
            if($su=="Transport-r") { $su="transport-refrigeration";}
            if($su=="Industrial-r") { $su="industrial-refrigeration";}
            if($su=="Commercial-r") { $su="commercial-refrigeration";}
            $term_id2 = term_exists( $su, "type", $term_id );

          if($term_id2) {
              wp_set_post_terms( $id, $term_id2, "type" ,true);
              }
          }
          /* INSERT TAGS */
          //true append all terms / false delete the previous //

          $ma=array_filter(explode("\n",$manufacturer));
          $ap=array_filter(explode("\n",$application));
          $re=array_filter(explode("\n",$refrigerant));
          $te=array_filter(explode("\n",$tech));
          $tags=array_filter(explode("\n",$tags));


          foreach($ma as $m) {
            $m=ucwords($m);
            wp_set_post_terms($id,$m,"manufacturer",true);
          }
          foreach($ap as $a) {
            $a=ucwords($a);
            wp_set_post_terms($id,$a,"application",true);
          }
          foreach($re as $r) {
            $r=ucfirst($r);
            wp_set_post_terms($id,$r,"refrigerant",true);
          }
          foreach($te as $t) {
          $t=ucwords($t);
           wp_set_post_terms($id,$t,"technology-type",true);
          }
          wp_set_post_terms($id,$country,"country");

          foreach($tags as $tag) {
            wp_set_post_terms($id,$tag,"post_tag",true);
          }


          add_post_meta($id, "source", $source, true);
          add_post_meta($id, "website", $web, true);
          add_post_meta($id, "energy_efficency", $energy,true);
    }

 $data=time();
 return $data;
}
?>

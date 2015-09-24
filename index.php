<?php
error_reporting(E_ALL);
include "Classes/phpQuery.php";
include "Classes/PHPExcel.php";
include "conf.php";

include "template/view.php";

date_default_timezone_set('UTC');
set_time_limit(0);

$url = 'http://artvulmarket.ru/pryazha-po-proizvoditelyam.html';

$root = parse_url($url, PHP_URL_SCHEME);
$root .= '://';
$root .= parse_url($url, PHP_URL_HOST);

phpQuery::newDocumentFileHTML($url);

$options_xls = PHPExcel_IOFactory::load(conf::$option_file);
$products_xls = PHPExcel_IOFactory::load(conf::$product_file);

//$options_xls->setActiveSheetIndex(0)->setCellValue('A1', 'Hello');

$l1 = pq('.js_list_x')->find('.div_promo_name>a');
$l1_cnt = count($l1);
$product_index = conf::$product_start_x;
$opid = conf::$option_options_start_x;
$opvolid = conf::$option_values_start_x;

foreach ($l1 as $i1 => $link) {

    if ($i1 == conf::$entry_limit) break;

    $href = pq($link)->attr('href');
    $manufacturer = pq($link)->text();

    phpQuery::newDocumentFileHTML($root . $href);

    $l2 = pq('.js_man_list')->find('.js_man_box');
    $l2_cnt = count($l2);

    foreach ($l2 as $i2 => $link) {

        if ($i2 == conf::$entry_limit) break;

        $href = pq($link)->find('.div_promo_name>a')->attr('href');
        $name = pq($link)->find('.div_promo_name>a')->text();

        $option_name = conf::$option_name_prefix.$name;

        $product_price = intval(pq($link)->find('.div_promo_price')->text());

        foreach (conf::$name_exclude as $exclude) {
            $name = str_replace($exclude, '', $name);
        }

        $name = trim($name);

        phpQuery::newDocumentFileHTML($root . $href);

        $main_image = pq('.js_man_sdesc_img>img')->attr('src');
        $product_image = '';

        if ($main_image !== '') {
            $fileext = pathinfo($main_image);//['extension'];
            $fileext = $fileext['extension'];
            $product_image = conf::$product_image_output_subdir . str_replace(' ', '_', $name) . '.' . $fileext;

            file_put_contents(conf::$product_image_output . $product_image, file_get_contents($main_image));
        }

        $p_desc =  pq('.js_man_sdesc_txt')->html();

        $product_length = intval(parseDesc($p_desc, 'Длина нити:'));
        $product_weight = intval(parseDesc($p_desc, 'Вес мотка:'));

        $products_xls->setActiveSheetIndex(conf::$product_sheet)
            ->setCellValue(conf::$product_col_id . $product_index, $product_index)
            ->setCellValue(conf::$product_col_name . $product_index, $name)
            ->setCellValue(conf::$product_col_cat . $product_index, conf::$product_cat)
            ->setCellValue(conf::$product_col_quantity . $product_index, conf::$product_quantity)
            ->setCellValue(conf::$product_col_model . $product_index, $name)
            ->setCellValue(conf::$product_col_manufacturer . $product_index, $manufacturer)
            ->setCellValue(conf::$product_col_image . $product_index, $product_image)
            ->setCellValue(conf::$product_col_shipping . $product_index, conf::$product_shipping)
            ->setCellValue(conf::$product_col_price . $product_index, $product_price)
            ->setCellValue(conf::$product_col_date_a . $product_index, date('Y-m-d h:m:s')) //2015-09-17 12:42:45
            ->setCellValue(conf::$product_col_date_m . $product_index, date('Y-m-d h:m:s')) //2015-09-17 12:42:45
            ->setCellValue(conf::$product_col_date_av . $product_index, date('Y-m-d')) //2015-09-17
            ->setCellValue(conf::$product_col_weight . $product_index, $product_weight)
            ->setCellValue(conf::$product_col_weight_unit . $product_index, conf::$product_w_unit)
            ->setCellValue(conf::$product_col_length . $product_index, $product_length)
            ->setCellValue(conf::$product_col_length_unit . $product_index, conf::$product_l_unit)
            ->setCellValue(conf::$product_col_status . $product_index, conf::$product_status)
            ->setCellValue(conf::$product_col_meta_t . $product_index, $name)
            ->setCellValue(conf::$product_col_meta_d . $product_index, $name)
            ->setCellValue(conf::$product_col_meta_k . $product_index, $name)
            ->setCellValue(conf::$product_col_stock_status . $product_index, conf::$product_stock_status)
            ->setCellValue(conf::$product_col_sort_order . $product_index, conf::$product_sort_order)
            ->setCellValue(conf::$product_col_substract . $product_index, conf::$product_substract)
            ->setCellValue(conf::$product_col_minimum . $product_index, conf::$product_minimum);

        $products_xls->setActiveSheetIndex(conf::$product_options_sheet)
            ->setCellValue(conf::$product_options_col_id . $product_index, $product_index)
            ->setCellValue(conf::$product_options_col_option . $product_index, $option_name)
            ->setCellValue(conf::$product_options_col_required . $product_index, conf::$product_options_required);


        $options_xls->setActiveSheetIndex(conf::$option_options_sheet)
            ->setCellValue(conf::$option_options_col_id . $opid, $opid)
            ->setCellValue(conf::$option_options_col_name . $opid, $option_name);

        $l3 = pq('.js_man_list')->find('.js_man_box');

        $items = [];
        foreach ($l3 as $i3 => $item) {
            $items[$i3]['name'] = pq($item)->find('.div_promo_name')->text();
            $items[$i3]['num'] = pq($item)->find('.div_promo_desc')->text();
            $items[$i3]['image'] = pq($item)->find('img.jshop_img')->attr('src');

        }

        if (count($items) > 1) {
            $str1 = $items[0]['name'];
            $str2 = $items[1]['name'];

            $len = mb_strlen($str2);
            for ($i = $len; $i > 0; $i--) {
                if (mb_stristr($str1, mb_substr($str2, 0, $i)) == $str1) {
                    if (($str2[$i - 1] == ' ') || ($i == $len)) {
                        $replace = mb_substr($str2, 0, $i);
                        break;
                    }
                }
            }
        } else {
            $replace = '';
        }

        $m_desc = [];

        foreach ($items as $item) {
            $option_value = str_replace($replace, '', $item['name']);
            $option_value = trim($option_value . ' ' . $item['num']);

            $img = $item['image'];

            if ($img !== '') {
                $fileext = pathinfo($img);//['extension'];
                $fileext = $fileext['extension'];
                $option_image = conf::$option_image_output_subdir . str_replace(' ', '_', $option_value) . '.' . $fileext;
                file_put_contents(conf::$option_image_output . $option_image, file_get_contents($img));
            }
            $m_desc[] = ['name' => $option_value, 'image' => $option_image];

            $products_xls->setActiveSheetIndex(conf::$product_option_values_sheet)
                ->setCellValue(conf::$product_option_values_col_id . $opvolid, $product_index)
                ->setCellValue(conf::$product_option_values_col_option . $opvolid, $option_name)
                ->setCellValue(conf::$product_option_values_col_option_value . $opvolid, $option_value)
                ->setCellValue(conf::$product_option_values_col_quantity . $opvolid, conf::$product_option_values_quantity)
                ->setCellValue(conf::$product_option_values_col_substract . $opvolid, conf::$product_option_values_substract)
                ->setCellValue(conf::$product_option_values_col_price . $opvolid, $product_price)
                ->setCellValue(conf::$product_option_values_col_price_prefix . $opvolid, conf::$product_option_values_price_prefix)
                ->setCellValue(conf::$product_option_values_col_points . $opvolid, conf::$product_option_values_points)
                ->setCellValue(conf::$product_option_values_col_points_prefix . $opvolid, conf::$product_option_values_points_prefix)
                ->setCellValue(conf::$product_option_values_col_weight . $opvolid, $product_weight)
                ->setCellValue(conf::$product_option_values_col_weight_prefix . $opvolid, conf::$product_option_values_weight_prefix);

            $options_xls->setActiveSheetIndex(conf::$option_values_sheet)
                ->setCellValue(conf::$option_values_col_id . $opvolid, $opvolid)
                ->setCellValue(conf::$option_values_col_option_id . $opvolid, $opid)
                //->setCellValue(conf::$option_values_col_image . $opvolid, $item['image'])  [10:54:07] ms. Sinister: картинки в опции не надо
                ->setCellValue(conf::$option_values_col_name . $opvolid, $option_value);

            $opvolid++;
        }

        ob_start();
        include('template/description.php');

        $f_desc = ob_get_contents();
        ob_end_clean();

        $products_xls->setActiveSheetIndex(conf::$product_sheet)
            ->setCellValue(conf::$product_col_desc . $product_index, $f_desc);

        $product_index++;
        $opid++;
        flush();
        ob_flush();
    }

}

$opt_save = PHPExcel_IOFactory::createWriter($options_xls, 'Excel5');
$opt_save->save(conf::$option_output);

$prod_save = PHPExcel_IOFactory::createWriter($products_xls, 'Excel5');
$prod_save->save(conf::$product_output);

$output_file = str_replace('{date}', date('d-m-y-h-m'), conf::$output_file);

$output = new ZipArchive();
$output->open($output_file, ZipArchive::CREATE);
$output->addFile(conf::$option_output);
$output->addFile(conf::$product_output);

$output->addGlob(conf::$product_image_output . conf::$product_image_output_subdir.'*.*');

$output->close();

//$link = '<a href="'.$output_file.'">'.$output_file.'</a>';

echo '<script>appendLink("'.$output_file.'")</script>';

function parseDesc($desc, $param) {
    $first = mb_strpos($desc, $param) + mb_strlen($param) + 1;
    $len = mb_strpos($desc, '<br>', $first);

    $len = $len - $first;
    return mb_substr($desc, $first, $len);
}


?>



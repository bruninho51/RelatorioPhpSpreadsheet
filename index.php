<?php 

    require_once 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
    use PhpOffice\PhpSpreadsheet\Style\Border;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    use PhpOffice\PhpSpreadsheet\Style\Alignment;
    
    function debug ($array) 
    {
        echo '<pre>';
        print_r($array);
        echo '</pre>';
    }

    if (isset($_POST['obterRelatorio'])) {
        
        $nomePlanilha = uniqid("qtde_Falta_").'.xlsx';

        $styleBorder = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => array('rgb' => '000000'),
                ),
            ),
        );

        $styleHeader = array(
            'font'  => array(
                'bold'  => false,
                'color' => array('rgb' => '000000'),
                'size'  => 12,
                'name'  => 'Calibre'
            ),
            'alignment' => array(
                'horizontal' => Alignment::HORIZONTAL_CENTER
            ),
            'fill' => array(
                'fillType' => Fill::FILL_GRADIENT_LINEAR,
                'color' => array('rgb' => '00AAFF')

            ),
            'borders' => $styleBorder['borders']
        );

        $styleSecondHeader = array(
            'font' => $styleHeader['font'],
            'fill' => array(
                'fillType' => Fill::FILL_GRADIENT_LINEAR,
                'color' => array('rgb' => 'BEBEBE')
            ),
            'borders' => $styleHeader['borders']
        );

        $dadosProd = array();
        $dadosEstab = array();
        
        $sqlProdutos = <<<HEREDOC
            SELECT DISTINCT P.* FROM produto P
                   JOIN estabProd EP ON EP.idProduto = P.idProduto ORDER BY nome ASC
        
HEREDOC;
        $sqlEstab = <<<HEREDOC
            SELECT DISTINCT E.* FROM estab E
                   JOIN estabProd EP ON EP.idEstab = E.idEstab ORDER BY nome ASC
        
HEREDOC;
        
        $sqlQtdeFaltaEstab = <<<HEREDOC
            SELECT DISTINCT
                    E.idEstab,
                    COALESCE(EPQF.qtde, 0),
                    COALESCE(EPQF.falta, 0)
            FROM estab E
            JOIN estabProd EP ON EP.idEstab = E.idEstab
            LEFT JOIN EstabProdQtdeFalta EPQF ON E.idEstab = EPQF.idEstab AND EPQF.idProduto = ?
            ORDER BY E.nome ASC
HEREDOC;
        
        $mysqli = new Mysqli("localhost","root","","phpRelatorio")or die("Erro de conexão");
        $resProd = $mysqli->query($sqlProdutos);
        $resEstab = $mysqli->query($sqlEstab);
        
        while ($reg = $resProd->fetch_assoc()) {
            array_push($dadosProd, $reg);
        }
        while ($reg = $resEstab->fetch_assoc()) {
            array_push($dadosEstab, $reg);
        }
        
        //CRIA A PLANILHA
        $excel = new Spreadsheet();
        
        //ATRIBUI PROPRIEDADES À PLANILHA
        $prop = $excel->getProperties();
        $prop->setTitle("Qtde e Faltas de Produtos");
        $prop->setDescription("Planilha de testes para o aprendizado da bilioteca PhpSpreadsheet.");

        //PEGA A PLANILHA ATIVA E ATRIBUI UM TÍTULO A ELA
        $planilha1 = $excel->getActiveSheet();
        $planilha1->setTitle("Qtde_Falta");
        
        for ($i=0,$col=2;$i<count($dadosProd);$i++,$col+=2) {

            //PEGA AS LETRAS DAS COLUNAS..
            $letCol = Coordinate::stringFromColumnIndex($col);
            $letProxCol = Coordinate::stringFromColumnIndex($col+1);
            $planilha1->mergeCells("{$letCol}1:{$letProxCol}1"); //MESCLA AS CÉLULAS
            
            //COLOCA O PRODUTO E PERSONALIZA A CÉLULA
            $planilha1->setCellValueByColumnAndRow($col, 1, $dadosProd[$i]['nome']);
            $planilha1->getStyle("{$letCol}1:{$letProxCol}1")->applyFromArray($styleHeader);

            //APROVEITA O LOOP PARA COLOCAR CABEÇARIO SECUNDARIO
            $planilha1->setCellValueByColumnAndRow($col, 2, "Qtde");
            $planilha1->setCellValueByColumnAndRow($col+1, 2, "Falta");
            $planilha1->getStyle("{$letCol}2")->applyFromArray($styleSecondHeader);
            $planilha1->getStyle("{$letProxCol}2")->applyFromArray($styleSecondHeader);
            
        }
        
        for ($i=0,$line=3;$i<count($dadosEstab);$i++) {
            $planilha1->getStyle('A'.$line)->applyFromArray($styleHeader);
            $planilha1->setCellValueByColumnAndRow(1, $line++, $dadosEstab[$i]['nome']);
        }
        $planilha1->getColumnDimension('A')->setAutoSize(true);

        $l=3; $c=2;
        foreach ($dadosProd as $prod) {
            $stmt = $mysqli->prepare($sqlQtdeFaltaEstab);
            $stmt->bind_param("s", $prod['idProduto']);
            $stmt->execute();
            $stmt->bind_result($idEstab, $qtde, $falta);
            
            while ($qtdeFalta = $stmt->fetch()) {
                $planilha1->getStyle(Coordinate::stringFromColumnIndex($c).$l)->applyFromArray($styleBorder);
                $planilha1->getStyle(Coordinate::stringFromColumnIndex($c+1).$l)->applyFromArray($styleBorder);

                $planilha1->setCellValueByColumnAndRow($c++,$l, $qtde);
                $planilha1->setCellValueByColumnAndRow($c, $l, $falta);
                $c--;$l++;
            }
            $c+=2;
            $l=3;
        }
        
        
        $writer = new Xlsx($excel);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$nomePlanilha.'"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
    }
    
?>

<!DOCTYPE html>
<html>
    <head>
        <meta charset='utf-8'>
        <title>Relatório</title>
    </head>
    <body>
        <form method="POST">
            <input type="hidden" name="obterRelatorio" value="relatorio">
            <button>Baixar Relatório</button>
        </form>
    </body>
</html>
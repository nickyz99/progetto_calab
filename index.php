<?php
// index.php

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Html as HtmlWriter;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// --- cleanup old exports (>60s) ---
function cleanupExports(string $dir, int $ttl = 60): void {
    if (!is_dir($dir)) {
        if (!mkdir($dir, 0755, true)) {
            error_log("Failed to create directory: $dir");
            return;
        }
    }
    if (!is_writable($dir)) {
        error_log("Directory not writable: $dir");
        return;
    }

    foreach (glob("$dir/*.xlsx") as $f) {
        if (time() - filemtime($f) > $ttl) {
            @unlink($f);
        }
    }
}
cleanupExports(__DIR__ . '/exports', 60);

// --- SQLite DB setup ---
$db = new PDO('sqlite:' . __DIR__ . '/data.db');
$db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$db->exec("CREATE TABLE IF NOT EXISTS products (
  id    INTEGER PRIMARY KEY AUTOINCREMENT,
  name  TEXT NOT NULL,
  price REAL NOT NULL
)");

// --- Router & Actions ---
$view = $_GET['view'] ?? 'products';

// Delete product
if ($view === 'products' && isset($_GET['delete_id'])) {
    $db->prepare("DELETE FROM products WHERE id = ?")
       ->execute([ intval($_GET['delete_id']) ]);
    header('Location:?view=products');
    exit;
}

// Edit/Add product
$editId = ($view === 'products' && isset($_GET['edit_id']))
        ? intval($_GET['edit_id'])
        : null;
if ($_SERVER['REQUEST_METHOD'] === 'POST' && $view === 'products') {
    if (!empty($_POST['edit_id'])) {
        $db->prepare("UPDATE products SET name = ?, price = ? WHERE id = ?")
           ->execute([ $_POST['name'], floatval($_POST['price']), intval($_POST['edit_id']) ]);
    } else {
        if (!empty($_POST['name'])) {
            $db->prepare("INSERT INTO products(name,price) VALUES(?,?)")
               ->execute([ $_POST['name'], floatval($_POST['price']) ]);
        }
    }
    header('Location:?view=products');
    exit;
}

// Download XLSX
if ($view === 'download' && $_SERVER['REQUEST_METHOD'] === 'POST') {
    $file_path_to_download = null;

    if (isset($_POST['preview_data'])) {
        // Generazione del file XLSX dai dati della preview modificati
        $cliente = $_POST['cliente_preview'] ?? 'Cliente';
        $parsed_preview_data = json_decode($_POST['preview_data'], true);

        $all_entries_final = [];
        if (is_array($parsed_preview_data)) {
            $current_date_entries = [];
            $current_date = '';
            foreach ($parsed_preview_data as $row_data) {
                if (isset($row_data['type']) && $row_data['type'] === 'date_label') {
                    if (!empty($current_date_entries)) {
                        $all_entries_final[] = ['date' => $current_date, 'entries' => $current_date_entries];
                        $current_date_entries = [];
                    }
                    $current_date = $row_data['date'];
                } elseif (isset($row_data['type']) && $row_data['type'] === 'product_row') {
                    $current_date_entries[] = [
                        'c'     => $row_data['colli'],
                        'name'  => $row_data['product_name'],
                        'kg'    => $row_data['kg'],
                        'price' => $row_data['price'],
                        'amount'=> $row_data['amount']
                    ];
                }
            }
            if (!empty($current_date_entries)) {
                $all_entries_final[] = ['date' => $current_date, 'entries' => $current_date_entries];
            }
        }

        $ss = new Spreadsheet();
        $sh = $ss->getActiveSheet();

        // Intestazione spostata su C1:G1
        $sh->mergeCells('C1:G1');
        $sh->setCellValue('C1', "NOTA DI VENDITA | $cliente");
        $sty = $sh->getStyle('C1');
        $sty->getFont()->setBold(true)->setSize(12)->setName('Arial');
        $sty->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Header su colonne C–G
        $hdr = ['C2'=>'C','D2'=>'PRODOTTO','E2'=>'KG','F2'=>'PREZZO','G2'=>'IMPORTO'];
        foreach ($hdr as $cell => $txt) {
          $sh->setCellValue($cell, $txt);
          $s = $sh->getStyle($cell);
          $s->getFont()->setBold(true);
          $s->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        // Larghezze colonne C–G
        foreach (['C'=>5,'D'=>20,'E'=>6,'F'=>12,'G'=>14] as $col=>$w) {
          $sh->getColumnDimension($col)->setWidth($w);
        }

        $r = 3;
        $max_data_rows = 22;
        $current_data_rows_in_excel = 0;

        $sum_start_row = 0;
        $sum_end_row = 0;

        foreach($all_entries_final as $date_section_data) {
            if ($current_data_rows_in_excel >= $max_data_rows) break;

            // Riga separatrice
            if ($current_data_rows_in_excel > 0) {
                $sh->getStyle("C{$r}:G{$r}")
                   ->applyFromArray([
                     'borders'=>[
                       'top'=>[
                         'borderStyle'=>Border::BORDER_DOUBLE,
                         'color'=>['argb'=>'FF000000']
                       ]
                     ]
                   ]);
                $r++;
                $current_data_rows_in_excel++;
                if ($current_data_rows_in_excel >= $max_data_rows) break;
            }

            $current_date_str = $date_section_data['date'];
            $display_date = '';
            if (!empty($current_date_str)) {
                $dt = \DateTime::createFromFormat('d/m/Y', $current_date_str);
                if ($dt !== false) {
                    $display_date = $dt->format('Y-m-d');
                } else {
                    $timestamp = strtotime($current_date_str);
                    if ($timestamp !== false) {
                        $display_date = date('Y-m-d', $timestamp);
                    }
                }
            }
            $sh->mergeCells("C{$r}:G{$r}");
            $sh->setCellValue("C{$r}", "Data: $display_date");
            $r++;
            $current_data_rows_in_excel++;
            if ($sum_start_row === 0) $sum_start_row = $r;

            // Righe prodotto su C–G
            foreach ($date_section_data['entries'] as $e) {
              if ($current_data_rows_in_excel >= $max_data_rows) break 2;
              $sh->setCellValue("C{$r}", $e['c'])
                 ->setCellValue("D{$r}", $e['name'])
                 ->setCellValue("E{$r}", $e['kg'])
                 ->setCellValue("F{$r}", $e['price'])
                 ->setCellValue("G{$r}", $e['amount']);
              $sh->getStyle("F{$r}")
                 ->getNumberFormat()
                 ->setFormatCode('"€"#,##0.00');
              $sh->getStyle("G{$r}")
                 ->getNumberFormat()
                 ->setFormatCode('"€"#,##0.00');
              $r++;
              $current_data_rows_in_excel++;
            }
        }

        $sum_end_row = $r - 1;

        // Riempi zeri su G fino a max_data_rows
        for (; $current_data_rows_in_excel < $max_data_rows; $current_data_rows_in_excel++) {
            $sh->setCellValue("G{$r}", 0);
            $sh->getStyle("G{$r}")
                ->getNumberFormat()
                ->setFormatCode('"€"#,##0.00');
            $r++;
        }

        $tot = $r;
        $sh->mergeCells("C{$tot}:F{$tot}");
        $sh->setCellValue("C{$tot}", "TOTALE");
        $formula = ($sum_start_row > 0 && $sum_end_row >= $sum_start_row)
                 ? "=SUM(G{$sum_start_row}:G{$sum_end_row})"
                 : "0";
        $sh->setCellValue("G{$tot}", $formula);
        $sh->getStyle("C{$tot}:G{$tot}")
           ->getFont()->setBold(true);
        $sh->getStyle("G{$tot}")
           ->getNumberFormat()
           ->setFormatCode('"€"#,##0.00');
        
        // <<< Inizio inserimento: applica bordo nero a tutto il template C2:G{$tot} >>>
        // Definisci lo stile bordo nero sottile
        $blackBorder = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color'       => ['argb' => 'FF000000'],
                ],
            ],
        ];
        // Applica il bordo a tutte le celle del template (header, dati e totale)
        $sh->getStyle("C2:G{$tot}")
           ->applyFromArray($blackBorder);
        // <<< Fine inserimento >>>
        
        // Salvataggio file
        if (!is_dir(__DIR__.'/exports')) mkdir(__DIR__.'/exports', 0755, true);
        $fname = "$cliente.xlsx";
        $file_path_to_download = __DIR__."/exports/$fname";
        (new Xlsx($ss))->save($file_path_to_download);

    } elseif (!empty($_POST['file'])) {
        $fname = basename($_POST['file']);
        $file_path_to_download = __DIR__ . "/exports/$fname";
    }

    if ($file_path_to_download && is_file($file_path_to_download)) {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment; filename=\"" . basename($file_path_to_download) . "\"");
        header('Content-Length: ' . filesize($file_path_to_download));
        readfile($file_path_to_download);
        exit;
    } else {
        error_log("File not found for download or path not set: " . ($file_path_to_download ?? 'NULL'));
        echo "Errore: File non trovato per il download.";
        exit;
    }
}

$products = $db->query("SELECT * FROM products")->fetchAll(PDO::FETCH_ASSOC);
?>
<!DOCTYPE html>
<html lang="it">
<!-- resto del template HTML invariato… -->

<head>
  <meta charset="UTF-8">
  <title>Gestione Vendite</title>
  <style>
    body { font-family:Arial,sans-serif; margin:2rem; }
    nav a { margin-right:1rem; text-decoration:none; font-weight:bold; }
    nav a.active { text-decoration:underline; }
    table { border-collapse:collapse; width:100%; margin-top:1rem; }
    th, td { border:1px solid #333; padding:0.5rem; text-align:center; }
    input, select { padding:0.3rem; width:100%; box-sizing:border-box; }
    .btn { padding:0.5rem 1rem; background:#007acc; color:#fff; border:none; cursor:pointer; }
    .btn:hover { background:#005fa3; }
    .editable-cell { cursor:text; }
    .date-section { margin-bottom: 1.5rem; padding: 1rem; border: 1px solid #ccc; background-color: #f9f9f9; }
    .date-section h4 { margin-top: 0; }
    /* Layout improvements */
    .form-group { margin-bottom: 1rem; }
    .form-group label { display: block; margin-bottom: 0.3rem; font-weight: bold; }
    .fieldset-container { border: 1px solid #ccc; padding: 1.5rem; margin-bottom: 2rem; border-radius: 5px; background-color: #fff; }
    .fieldset-container legend { font-weight: bold; font-size: 1.2em; padding: 0 0.5rem; color: #007acc; }
    .btn-group { margin-top: 1.5rem; }
  </style>
  <script>
    document.addEventListener('DOMContentLoaded', function(){
      var cb = document.getElementById('multi_date');

      if(cb) {
          var multiDateSection = document.getElementById('multi_date_section'),
              addDateBtn = document.getElementById('add_date_btn'),
              dateSectionsContainer = document.getElementById('date_sections_container'),
              firstDateSection = document.getElementById('date_section_0');

          function toggleMultiDateSection(){
              if (cb.checked) {
                  multiDateSection.style.display = 'block';
                  firstDateSection.style.display = 'none';
                  if (dateSectionsContainer.children.length === 0) {
                      // Aggiungi la prima sezione data con la data odierna se non ce ne sono
                      addDateSection(0, new Date().toISOString().slice(0,10));
                  }
              } else {
                  multiDateSection.style.display = 'none';
                  firstDateSection.style.display = 'block';
                  dateSectionsContainer.innerHTML = ''; // Rimuovi tutte le sezioni extra
                  // Reset della singola data se necessario
                  const singleDateInput = document.getElementById('date1');
                  if (singleDateInput) {
                      singleDateInput.value = new Date().toISOString().slice(0,10);
                  }
                  // Reset dei campi kg e colli per la singola data
                  Array.from(firstDateSection.querySelectorAll('input[type="number"]')).forEach(input => input.value = 0);
              }
          }

          let dateCounter = 0; // Inizializza a 0, sarà incrementato o usato come 0 per la prima sezione
          function addDateSection(initialCounter = null, initialDate = '') {
              const currentCounter = initialCounter !== null ? initialCounter : ++dateCounter;

              const newSection = document.createElement('div');
              newSection.classList.add('date-section');
              newSection.setAttribute('data-date-index', currentCounter);
              newSection.innerHTML = `
                  <h4>Data ${currentCounter + 1}</h4>
                  <div class="form-group">
                    <label>Data:</label>
                    <input type="date" name="dates[${currentCounter}][date]" value="${initialDate}" required>
                  </div>
                  <table>
                      <thead>
                          <tr><th>Colli</th><th>Prodotto</th><th>KG</th></tr>
                      </thead>
                      <tbody>
                          <?php foreach($products as $i=>$p): ?>
                          <tr>
                              <td><input name="dates[${currentCounter}][colli][<?=$i?>]" type="number" min="0" value="0"></td>
                              <td><?=htmlspecialchars($p['name'])?></td>
                              <td><input name="dates[${currentCounter}][kg][<?=$i?>]" type="number" step="0.01" min="0" value="0"></td>
                          </tr>
                          <?php endforeach;?>
                      </tbody>
                  </table>
                  <button type="button" class="btn remove-date-btn" style="background-color: #dc3545; margin-top: 10px;">Rimuovi Data</button>
              `;
              dateSectionsContainer.appendChild(newSection);

              newSection.querySelector('.remove-date-btn').addEventListener('click', function() {
                  newSection.remove();
                  // Se la checkbox è spuntata e non ci sono più sezioni, ripristina la prima sezione.
                  if (cb.checked && dateSectionsContainer.children.length === 0) {
                      // Se tutte le sezioni sono state rimosse e la multi-data è ancora attiva, riaggiungi la prima
                      addDateSection(0, new Date().toISOString().slice(0,10));
                  }
              });
          }

          // Inizializza lo stato al caricamento della pagina
          toggleMultiDateSection();
          cb.addEventListener('change', toggleMultiDateSection);

          if (addDateBtn) {
              addDateBtn.addEventListener('click', function() {
                  addDateSection(null, new Date().toISOString().slice(0,10));
              });
          }

          // Se la checkbox non è spuntata al caricamento (es. al ritorno da POST con errore), assicurati che la singola data sia visibile.
          // Questo blocco è già gestito da toggleMultiDateSection() al DCL
      }
    });
  </script>
</head>
<body>

<nav>
  <a href="?view=products" class="<?= $view==='products'?'active':'' ?>">Prodotti</a>
  <a href="?view=template" class="<?= $view==='template'?'active':'' ?>">Genera Template</a>
</nav>

<?php if ($view === 'products'): ?>

  <h2>Gestione Prodotti</h2>

  <?php if ($editId):
    $st = $db->prepare("SELECT * FROM products WHERE id=?");
    $st->execute([$editId]);
    $p0 = $st->fetch(PDO::FETCH_ASSOC);
  ?>
    <form method="post">
      <input type="hidden" name="edit_id" value="<?=$p0['id']?>">
      <table>
        <tr><th>Nome</th><th>Prezzo (€)</th><th>Salva</th></tr>
        <tr>
          <td><input name="name" value="<?=htmlspecialchars($p0['name'])?>" required></td>
          <td><input name="price" type="number" step="0.01" value="<?=$p0['price']?>" required></td>
          <td><button class="btn">Salva</button></td>
        </tr>
      </table>
    </form><hr>
  <?php endif; ?>

  <form method="post">
    <table>
      <tr><th>Nome Prodotto</th><th>Prezzo (€)</th><th>Aggiungi</th></tr>
      <tr>
        <td><input name="name" placeholder="Nome prodotto" required></td>
        <td><input name="price" type="number" step="0.01" placeholder="€" required></td>
        <td><button class="btn">Aggiungi</button></td>
      </tr>
    </table>
  </form>

  <?php if ($products): ?>
    <table>
      <tr><th>ID</th><th>Nome</th><th>Prezzo (€)</th><th>Modifica</th><th>Elimina</th></tr>
      <?php foreach($products as $p): ?>
      <tr>
        <td><?=$p['id']?></td>
        <td><?=htmlspecialchars($p['name'])?></td>
        <td><?=number_format($p['price'],2,',','.')?></td>
        <td><a href="?view=products&edit_id=<?=$p['id']?>" class="btn">✏️</a></td>
        <td><a href="?view=products&delete_id=<?=$p['id']?>" class="btn" onclick="return confirm('Elimina?')">🗑️</a></td>
      </tr>
      <?php endforeach;?>
    </table>
  <?php else: ?>
    <p>Nessun prodotto presente.</p>
  <?php endif; ?>

<?php elseif ($view === 'template'): ?>

  <?php
    $preview_data_received = isset($_POST['preview_data']);
    $all_entries_final = [];
    $cliente = '';

    if ($preview_data_received) {
        // Se i dati vengono dalla preview, usali direttamente
        $cliente = $_POST['cliente_preview'] ?? 'Cliente';
        $parsed_preview_data = json_decode($_POST['preview_data'], true);

        if (is_array($parsed_preview_data)) {
            $current_date_entries = [];
            $current_date = '';
            foreach ($parsed_preview_data as $row_data) {
                if (isset($row_data['type']) && $row_data['type'] === 'date_label') {
                    if (!empty($current_date_entries)) {
                        $all_entries_final[] = ['date' => $current_date, 'entries' => $current_date_entries];
                        $current_date_entries = [];
                    }
                    $current_date = $row_data['date'];
                } elseif (isset($row_data['type']) && $row_data['type'] === 'product_row') {
                    $current_date_entries[] = [
                        'c'     => $row_data['colli'],
                        'name'  => $row_data['product_name'],
                        'kg'    => $row_data['kg'],
                        'price' => $row_data['price']
                    ];
                }
            }
            if (!empty($current_date_entries)) {
                $all_entries_final[] = ['date' => $current_date, 'entries' => $current_date_entries];
            }
        }
    } else if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        // Se i dati vengono dal form iniziale, elaborali
        $cliente = preg_replace('/[^A-Za-z0-9_\-]/','_', substr($_POST['cliente'],0,50));

        if (isset($_POST['multi_date'])) {
            if (isset($_POST['dates']) && is_array($_POST['dates'])) {
                foreach ($_POST['dates'] as $date_index => $date_data) {
                    $current_date = $date_data['date'];
                    $date_entries = [];
                    foreach ($products as $p_index => $pr) {
                        $c = intval($date_data['colli'][$p_index] ?? 0);
                        $kg = floatval($date_data['kg'][$p_index] ?? 0);
                        if ($c || $kg) { // Aggiungi solo prodotti con colli o KG > 0
                            $date_entries[] = ['c'=>$c,'name'=>$pr['name'],'kg'=>$kg,'price'=>$pr['price']];
                        }
                    }
                    if (!empty($date_entries)) {
                        $all_entries_final[] = ['date' => $current_date, 'entries' => $date_entries];
                    }
                }
            }
        } else {
            $date1 = $_POST['date1'];
            $single_date_entries = [];
            foreach ($products as $i => $pr) {
              $c1 = intval($_POST["colli1_$i"]);
              $kg1 = floatval($_POST["kg1_$i"]);
              if ($c1 || $kg1) { // Aggiungi solo prodotti con colli o KG > 0
                  $single_date_entries[] = ['c'=>$c1,'name'=>$pr['name'],'kg'=>$kg1,'price'=>$pr['price']];
              }
            }
            if (!empty($single_date_entries)) {
                $all_entries_final[] = ['date' => $date1, 'entries' => $single_date_entries];
            }
        }
    }

    // Questa parte del codice viene eseguita solo se ci sono dati da mostrare in preview
    if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($all_entries_final)) :

        $ss = new Spreadsheet();
        $sh = $ss->getActiveSheet();

        // Intestazione Cliente
        $sh->mergeCells('A1:E1');
        $sh->setCellValue('A1', "NOTA DI VENDITA | $cliente");
        $sty = $sh->getStyle('A1');
        $sty->getFont()->setBold(true)->setSize(12)->setName('Arial');
        $sty->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Header della tabella
        $hdr = ['A2'=>'C','B2'=>'PRODOTTO','C2'=>'KG','D2'=>'PREZZO','E2'=>'IMPORTO'];
        foreach ($hdr as $cell => $txt) {
          $sh->setCellValue($cell, $txt);
          $s = $sh->getStyle($cell);
          $s->getFont()->setBold(true);
          $s->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        // Larghezza colonne
        foreach (['A'=>5,'B'=>20,'C'=>6,'D'=>12,'E'=>14] as $col=>$w) {
          $sh->getColumnDimension($col)->setWidth($w);
        }

        $r = 3; // Inizia i dati dalla riga 3
        $max_data_rows = 22; // Righe massime di dati visualizzabili per pagina Excel
        $current_data_rows_in_excel = 0;

        $sum_start_row = 0;
        $sum_end_row = 0;

        foreach($all_entries_final as $date_section_data) {
            if ($current_data_rows_in_excel >= $max_data_rows) break;

            $current_date_str = $date_section_data['date'];
            $entries = $date_section_data['entries'];

            // Aggiungi una riga separatrice se non è la prima sezione data
            if ($current_data_rows_in_excel > 0) {
                $sh->getStyle("A{$r}:E{$r}")
                   ->applyFromArray([
                     'borders'=>[
                       'top'=>[
                         'borderStyle'=>Border::BORDER_DOUBLE,
                         'color'=>['argb'=>'FF000000']
                       ]
                     ]
                   ]);
                $r++;
                $current_data_rows_in_excel++;
                if ($current_data_rows_in_excel >= $max_data_rows) break;
            }

            // Inserisci la data
            $display_date = '';
            $timestamp = strtotime($current_date_str);
            // Assicurati che il timestamp sia valido E che la stringa originale non sia vuota
            if ($timestamp !== false && $current_date_str !== '') {
                $display_date = date('d/m/Y', $timestamp);
            }
            $sh->mergeCells("A{$r}:E{$r}");
            $sh->setCellValue("A{$r}", "Data: " . $display_date);
            $r++;
            $current_data_rows_in_excel++;
            if ($sum_start_row === 0) $sum_start_row = $r; // Imposta la riga di inizio per la somma

            // Inserisci i prodotti per questa data
            foreach ($entries as $e) {
              if ($current_data_rows_in_excel >= $max_data_rows) break 2;
              $sh->setCellValue("A{$r}", $e['c'])
                 ->setCellValue("B{$r}", $e['name'])
                 ->setCellValue("C{$r}", $e['kg'])
                 ->setCellValue("D{$r}", $e['price'])
                 ->setCellValue("E{$r}", $e['amount'] ?? ($e['kg'] * $e['price'])); // Se amount è presente dal JSON (preview), usalo. Altrimenti calcolalo.
              $sh->getStyle("D{$r}")->getNumberFormat()->setFormatCode('"€"#,##0.00');
              $sh->getStyle("E{$r}")->getNumberFormat()->setFormatCode('"€"#,##0.00');
              $r++;
              $current_data_rows_in_excel++;
            }
        }

        $sum_end_row = $r - 1; // La fine del range di somma è la riga precedente a quella attuale

        // Riempie le righe rimanenti con valori 0 nell'importo per mantenere la dimensione fissa del template
        for (; $current_data_rows_in_excel < $max_data_rows; $current_data_rows_in_excel++) {
            $sh->setCellValue("E{$r}", 0); // Inserisce 0 come valore numerico
            $sh->getStyle("E{$r}")->getNumberFormat()->setFormatCode('"€"#,##0.00');
            $r++;
        }

        // Riga Totale
        $tot = $r;
        $sh->mergeCells("A{$tot}:D{$tot}");
        $sh->setCellValue("A{$tot}", "TOTALE");
        // Calcola la somma solo se ci sono state righe di dati valide per la somma
        $sh->setCellValue("E{$tot}", ($sum_start_row > 0 && $sum_end_row >= $sum_start_row) ? "=SUM(E{$sum_start_row}:E{$sum_end_row})" : 0);
        $sh->getStyle("A{$tot}:E{$tot}")->getFont()->setBold(true);
        $sh->getStyle("E{$tot}")->getNumberFormat()->setFormatCode('"€"#,##0.00');

        // Salvataggio del file XLSX
        if (!is_dir(__DIR__.'/exports')) mkdir(__DIR__.'/exports', 0755, true);
        $fname = "$cliente.xlsx";
        $file_path_to_download = __DIR__."/exports/$fname";
        (new Xlsx($ss))->save($file_path_to_download);

        // Generazione dell'HTML per la preview
        $writer = new HtmlWriter($ss);
        $writer->setSheetIndex(0);
        $html = $writer->generateSheetData();

        // Manipolazione del DOM per rendere le celle editabili e aggiungere data-attributi
        libxml_use_internal_errors(true);
        $dom = new DOMDocument();
        $dom->loadHTML($html, LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD);
        libxml_clear_errors();

        $table = $dom->getElementsByTagName('table')->item(0);
        if ($table) {
            $table->setAttribute('id', 'preview');
            $rows = $table->getElementsByTagName('tr');
            $headerCells = []; $dataStart = null;

            // Trova gli indici delle colonne dagli header per sapere dove sono KG, Prezzo, Importo
            foreach ($rows as $idx => $row) {
                $cells = $row->getElementsByTagName('td'); // PHPSpreasheet genera td anche per gli header
                if ($cells->length > 0) {
                    if ($idx === 1) { // Seconda riga (indice 1) contiene gli header di colonna
                        foreach ($cells as $h => $c) {
                            $t = trim(strtoupper($c->textContent));
                            if (strpos($t,'C')===0)       $headerCells['colli']=$h;
                            if (strpos($t,'KG')===0)     $headerCells['kg']=$h;
                            if (strpos($t,'PREZZO')===0) $headerCells['price']=$h;
                            if (strpos($t,'IMPORTO')===0)$headerCells['amount']=$h;
                            if (strpos($t,'PRODOTTO')===0)$headerCells['product']=$h;
                        }
                    }
                    // Trova la prima riga di dati (non header, non riga data, non totale)
                    if ($dataStart===null && $idx>1 && strpos($cells->item(0)->textContent,'Data:')===false && trim($cells->item(0)->textContent) !== 'TOTALE') {
                        $dataStart=$idx;
                    }
                }
            }

            // Rendi le celle editabili e aggiungi data-attributi
            foreach ($rows as $idx => $row) {
                $cells = $row->getElementsByTagName('td');
                if ($cells->length>0) {
                    $isTotal=false;
                    // Controlla se è la riga del totale
                    foreach ($cells as $c) {
                        if (trim(strtoupper($c->textContent))==='TOTALE') {
                            $isTotal=true; break;
                        }
                    }
                    $isDateLabelRow = ($cells->length === 1 && strpos($cells->item(0)->textContent, 'Data:') === 0);

                    // Solo le righe di dati (non header, non data, non totale)
                    if ($idx>=$dataStart && !$isTotal && !$isDateLabelRow) {
                        foreach (['kg','price','amount'] as $type) {
                            if (isset($headerCells[$type]) && $cells->length > $headerCells[$type]) {
                                $c = $cells->item($headerCells[$type]);
                                $c->setAttribute('contenteditable','true');
                                $c->setAttribute('class','editable-cell');
                                $c->setAttribute('data-col-type',$type);
                            }
                        }
                        // Aggiungi data-col-type anche per Prodotto e Colli (non editabili, ma utili per JS)
                        if (isset($headerCells['product']) && $cells->length > $headerCells['product']) {
                            $cells->item($headerCells['product'])->setAttribute('data-col-type','product');
                        }
                        if (isset($headerCells['colli']) && $cells->length > $headerCells['colli']) {
                            $cells->item($headerCells['colli'])->setAttribute('data-col-type','colli');
                        }
                    }
                    // La cella del totale non è editabile
                    if ($isTotal) {
                        $last = $cells->item($cells->length-1);
                        $last->setAttribute('id','totalAmount');
                        $last->setAttribute('data-col-type','total');
                        $last->setAttribute('contenteditable','false');
                        $last->setAttribute('class','editable-cell'); // Mantiene la classe per lo stile
                    }
                }
            }
            $html = $dom->saveHTML();
        }

        echo '<h3>Preview (valori calcolati)</h3>';
        echo $html;
        ?>
        <form id="downloadForm" action="?view=download" method="post" style="display:none;" target="_blank">
            <input type="hidden" name="file" value="<?= htmlspecialchars($fname) ?>">
            <input type="hidden" name="preview_data" id="preview_data_input">
            <input type="hidden" name="cliente_preview" value="<?= htmlspecialchars($cliente) ?>">
        </form>
        <p><button class="btn" id="download_xlsx_btn">Download .xlsx</button></p>
        <script>
        document.addEventListener('DOMContentLoaded',function(){
          const tbl=document.getElementById('preview');
          const totCell=document.getElementById('totalAmount');
          const downloadBtn = document.getElementById('download_xlsx_btn');
          const downloadForm = document.getElementById('downloadForm');
          const previewDataInput = document.getElementById('preview_data_input');


          if(!tbl||!totCell) return;

          // Funzione per pulire il valore e convertirlo in float
          function clean(v){
            // Rimuove simbolo euro, caratteri non numerici (tranne virgola/punto per decimali) e sostituisce virgola con punto
            // Gestisce il caso in cui v sia null o undefined restituendo 0
            return parseFloat((v || '').replace('€', '').replace(/[^0-9,.-]/g,'').replace(',','.')) || 0;
          }

          // Funzione per formattare il valore in euro
          function fmt(v){
            // Assicura che il valore sia un numero prima di formattare
            const num = parseFloat(v);
            if (isNaN(num)) return '€ 0,00'; // Gestisce casi in cui il numero non è valido
            return num.toLocaleString('it-IT',{style:'currency',currency:'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2});
          }

          // Funzione per ricalcolare una singola riga
          function recalcRow(row, modifiedType){
            const kgCell    = row.querySelector('[data-col-type=kg]');
            const priceCell = row.querySelector('[data-col-type=price]');
            const amountCell= row.querySelector('[data-col-type=amount]');

            // Assicurati che le celle esistano prima di tentare di leggere textContent
            if (!kgCell || !priceCell || !amountCell) return;

            let kg    = clean(kgCell.textContent);
            let price = clean(priceCell.textContent);
            let amount= clean(amountCell.textContent);

            if (modifiedType === 'kg' || modifiedType === 'price') {
              amount = kg * price; // Se modifichi KG o Prezzo, ricalcola l'Importo
            } else if (modifiedType === 'amount') {
              // Se modifichi l'Importo, ricalcola i KG
            
              // L'importo (amount) è già quello inserito dall'utente, non lo ricalcoliamo
            }

            // Aggiorna i contenuti delle celle nella preview
            // I KG e il Prezzo mantengono la loro formattazione numerica standard, mentre l'Importo ha la valuta.
            kgCell.textContent     = kg.toLocaleString('it-IT', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            priceCell.textContent  = fmt(price);
            amountCell.textContent = fmt(amount);
          }

          // Funzione per ricalcolare il totale generale
          function recalcAll(){
            let sum = 0;
            Array.from(tbl.rows).forEach((row,i)=>{
              // Ottieni le celle da questa riga
              const cells = Array.from(row.querySelectorAll('td'));

              // Salta le righe che non sono righe di prodotto (header, data, totale)
              const isHeaderRow = cells.some(cell => cell.tagName === 'TH' || (cell.dataset && cell.dataset.colType === 'product' && cell.textContent === 'PRODOTTO')); // Più robusto per header
              const isDateLabelRow = cells.length === 1 && cells[0].textContent.startsWith('Data:');
              const isTotalRow = row.querySelector('#totalAmount');

              if(isHeaderRow || isDateLabelRow || isTotalRow) return;

              const amountCell = row.querySelector('[data-col-type=amount]');
              if (amountCell) {
                 sum += clean(amountCell.textContent);
              }
            });
            totCell.textContent = fmt(sum);
          }

          // Event listener per le modifiche nelle celle
          tbl.addEventListener('input',function(event){
            const target = event.target;
            // Assicurati che l'elemento sia una cella editabile (non il totale)
            if(target.classList.contains('editable-cell') && target.dataset.colType !== 'total'){
              const row = target.closest('tr');
              const type = target.dataset.colType;
              recalcRow(row, type); // Ricalcola la riga specifica
              recalcAll(); // Ricalcola il totale generale
            }
          });

          // Event listener per il pulsante di download
          downloadBtn.addEventListener('click', function() {
            const dataToSave = [];

            // Itera su tutte le righe della tabella di preview per raccogliere i dati aggiornati
            Array.from(tbl.rows).forEach((row) => {
                const cells = Array.from(row.querySelectorAll('td'));

                // Determina il tipo di riga
                const isHeaderRow = cells.some(cell => cell.tagName === 'TH' || (cell.dataset && cell.dataset.colType === 'product' && cell.textContent === 'PRODOTTO')); // Più robusto per header
                const isDateLabelRow = cells.length === 1 && cells[0].textContent.startsWith('Data:');
                const isTotalRow = row.querySelector('#totalAmount');

                if (isHeaderRow || isTotalRow) {
                    return; // Salta le righe di intestazione e la riga del totale
                } else if (isDateLabelRow) {
                    const date_text = cells[0].textContent.replace('Data: ', '');
                    dataToSave.push({ type: 'date_label', date: date_text });
                } else {
                    const colliCell = row.querySelector('[data-col-type=colli]');
                    const productNameCell = row.querySelector('[data-col-type=product]');
                    const kgCell = row.querySelector('[data-col-type=kg]');
                    const priceCell = row.querySelector('[data-col-type=price]');
                    const amountCell = row.querySelector('[data-col-type=amount]');

                    // Raccoglie i dati solo se la riga contiene tutte le celle rilevanti E ha valori significativi
                    // Ho aggiunto un controllo più robusto prima di accedere a textContent
                    if (colliCell && productNameCell && kgCell && priceCell && amountCell) {
                        const current_kg = clean(kgCell.textContent);
                        const current_price = clean(priceCell.textContent);
                        const current_amount = clean(amountCell.textContent);

                        // Includi la riga solo se KG o Importo sono maggiori di zero
                        if (current_kg > 0 || current_amount > 0) {
                            dataToSave.push({
                                type: 'product_row',
                                colli: clean(colliCell.textContent), // Assicurati che colliCell esista prima di leggerlo
                                product_name: productNameCell.textContent,
                                kg: current_kg,
                                price: current_price,
                                amount: current_amount
                            });
                        }
                    }
                }
            });

            // Inserisci i dati JSON nel campo nascosto del form e invia
            previewDataInput.value = JSON.stringify(dataToSave);
            downloadForm.submit();
          });

          recalcAll(); // Calcola il totale iniziale al caricamento della preview
        });
        </script>
        <?php exit; ?>
    <?php endif; // Fine if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($all_entries_final)) ?>

    <?php // Se non è stato inviato il form o non ci sono dati, mostra il form iniziale ?>
  <h2>Genera Template</h2>
  <form method="post">
    <div class="fieldset-container">
        <legend>Dati Cliente e Data</legend>
        <div class="form-group">
            <label for="cliente">Cliente:</label>
            <input id="cliente" name="cliente" required value="<?= htmlspecialchars($_POST['cliente'] ?? '') ?>">
        </div>

        <div id="date_section_0" class="form-group">
            <label for="date1">Data:</label>
            <input type="date" id="date1" name="date1" value="<?=htmlspecialchars($_POST['date1'] ?? date('Y-m-d'))?>" required>
            <table>
                <thead>
                    <tr><th>Colli</th><th>Prodotto</th><th>KG</th></tr>
                </thead>
                <tbody>
                    <?php foreach($products as $i=>$p): ?>
                    <tr>
                        <td><input name="colli1_<?=$i?>" type="number" min="0" value="<?= htmlspecialchars($_POST["colli1_$i"] ?? 0) ?>"></td>
                        <td><?=htmlspecialchars($p['name'])?></td>
                        <td><input name="kg1_<?=$i?>" type="number" step="0.01" min="0" value="<?= htmlspecialchars($_POST["kg1_$i"] ?? 0) ?>"></td>
                    </tr>
                    <?php endforeach;?>
                </tbody>
            </table>
        </div>

        <div class="form-group">
            <label><input type="checkbox" id="multi_date" name="multi_date" <?= isset($_POST['multi_date']) ? 'checked' : '' ?>> Abilita date multiple</label>
        </div>

        <div id="multi_date_section" style="display:none;" class="fieldset-container">
            <legend>Date Aggiuntive</legend>
            <div id="date_sections_container">
                <?php
                // Ricostruisce le sezioni multi-data in caso di ricaricamento pagina (es. per errore di validazione)
                if (isset($_POST['multi_date']) && isset($_POST['dates']) && is_array($_POST['dates'])) {
                    foreach ($_POST['dates'] as $date_index => $date_data) {
                        // Assicurati che la data sia sempre impostata, anche se $_POST['date'] fosse vuoto
                        $current_date_val = htmlspecialchars($date_data['date'] ?? date('Y-m-d'));
                        echo "<div class=\"date-section\" data-date-index=\"{$date_index}\">";
                        echo "<h4>Data " . ($date_index + 1) . "</h4>";
                        echo "<div class=\"form-group\">";
                        echo "<label>Data:</label>";
                        echo "<input type=\"date\" name=\"dates[{$date_index}][date]\" value=\"{$current_date_val}\" required>";
                        echo "</div>";
                        echo "<table>";
                        echo "<thead><tr><th>Colli</th><th>Prodotto</th><th>KG</th></tr></thead>";
                        echo "<tbody>";
                        foreach ($products as $p_index => $pr) {
                            $colli_val = htmlspecialchars($date_data['colli'][$p_index] ?? 0);
                            $kg_val = htmlspecialchars($date_data['kg'][$p_index] ?? 0);
                            echo "<tr>";
                            echo "<td><input name=\"dates[{$date_index}][colli][{$p_index}]\" type=\"number\" min=\"0\" value=\"{$colli_val}\"></td>";
                            echo "<td>" . htmlspecialchars($pr['name']) . "</td>";
                            echo "<td><input name=\"dates[{$date_index}][kg][{$p_index}]\" type=\"number\" step=\"0.01\" min=\"0\" value=\"{$kg_val}\"></td>";
                            echo "</tr>";
                        }
                        echo "</tbody>";
                        echo "</table>";
                        echo "<button type=\"button\" class=\"btn remove-date-btn\" style=\"background-color: #dc3545; margin-top: 10px;\">Rimuovi Data</button>";
                        echo "</div>";
                    }
                }
                ?>
            </div>
            <button type="button" id="add_date_btn" class="btn" style="margin-top: 1rem;">Aggiungi Data</button>
        </div>
    </div>

    <div class="btn-group">
        <button class="btn">Genera & Preview</button>
    </div>
  </form>
<?php endif; // Fine if ($view === 'template') ?>

</body>
</html>
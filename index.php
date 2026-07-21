<?php
/**
 * Dynamische Ergebnisliste
 * Findet alle HTML-Dateien und extrahiert die Namen aus zugehörigen XML-Dateien
 */

// Hole alle HTML-Dateien aus dem aktuellen Verzeichnis (außer index.php und tmp_results)
$currentDir = dirname(__FILE__);
$htmlFiles = array_filter(glob($currentDir . '/*.html'), function($file) {
    return strpos(basename($file), 'tmp_results') === false;
});
$htmlFiles = array_values($htmlFiles); // Neu indexieren
sort($htmlFiles);

// Funktion zum Extrahieren des Namens aus einer XML-Datei
function getNameFromXml($xmlPath) {
    if (!file_exists($xmlPath)) {
        return null;
    }
    
    try {
        $xml = simplexml_load_file($xmlPath);
        if ($xml === false) {
            return null;
        }
        
        // Das Root-Element sollte 'result' sein
        if ($xml->getName() === 'result') {
            $name = (string)$xml['name'];
            return !empty($name) ? $name : null;
        }
        
        return null;
    } catch (Exception $e) {
        return null;
    }
}

// Funktion zum Extrahieren des Datums aus einer XML-Datei
function getDateFromXml($xmlPath) {
    if (!file_exists($xmlPath)) {
        return null;
    }
    
    try {
        $xml = simplexml_load_file($xmlPath);
        if ($xml === false) {
            return null;
        }
        
        // Das Root-Element sollte 'result' sein
        if ($xml->getName() === 'result') {
            $date = (string)$xml['date'];
            return !empty($date) ? $date : null;
        }
        
        return null;
    } catch (Exception $e) {
        return null;
    }
}

// Erstelle die Liste der Einträge
$entries = [];
$latestDate = null;

foreach ($htmlFiles as $htmlFile) {
    $htmlFileName = basename($htmlFile);
    
    // Finde die zugehörige XML-Datei
    $xmlFile = str_replace('.html', '.xml', $htmlFile);
    
    if (file_exists($xmlFile)) {
        $name = getNameFromXml($xmlFile);
        $date = getDateFromXml($xmlFile);
        
        if ($name) {
            $entries[] = [
                'html' => $htmlFileName,
                'name' => $name,
                'date' => $date
            ];
            
            // Finde das neueste Datum
            if ($date) {
                // Konvertiere das Datum in ein Format, das vergleichbar ist
                // Format: "21.07.2026 16:34:21" -> DateTime
                try {
                    $dateTime = DateTime::createFromFormat('d.m.Y H:i:s', $date);
                    if ($dateTime && (!$latestDate || $dateTime > $latestDate)) {
                        $latestDate = $dateTime;
                    }
                } catch (Exception $e) {
                    // Ignorieren bei Fehler
                }
            }
        }
    }
}
?>
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gauschießen 2026</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #0a1628 0%, #162d42 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        
        .container {
            background: linear-gradient(to bottom, #0f1f2e, #0a1628);
            border-radius: 10px;
            box-shadow: 0 15px 50px rgba(0, 0, 0, 0.5);
            padding: 40px;
            max-width: 600px;
            width: 100%;
            border: 1px solid #1a3a4a;
        }
        
        h1 {
            color: #ffffff;
            margin-bottom: 30px;
            text-align: center;
            font-size: 28px;
            position: relative;
            padding-bottom: 15px;
        }
        
        h1::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 50%;
            transform: translateX(-50%);
            width: 60px;
            height: 3px;
            background: linear-gradient(to right, #80dd00, #1fa9a4);
            border-radius: 2px;
        }
        
        .list-container {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .list-item {
            background: linear-gradient(to right, rgba(31, 169, 164, 0.1), rgba(128, 221, 0, 0.05));
            border: 1px solid #1fa9a4;
            border-radius: 6px;
            transition: all 0.3s ease;
            overflow: hidden;
        }
        
        .list-item:hover {
            background: linear-gradient(to right, rgba(31, 169, 164, 0.2), rgba(128, 221, 0, 0.1));
            border-color: #80dd00;
            transform: translateX(5px);
            box-shadow: 0 5px 20px rgba(31, 169, 164, 0.2);
        }
        
        .list-item a {
            display: block;
            padding: 15px 20px;
            color: #ffffff;
            text-decoration: none;
            font-weight: 500;
            font-size: 16px;
            transition: color 0.3s ease;
        }
        
        .list-item a:hover {
            color: #80dd00;
        }
        
        .info {
            text-align: center;
            color: #7fa3b3;
            padding: 20px;
            font-style: italic;
            font-size: 14px;
            margin-top: 15px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Ergebnislisten Gauschießen 2026</h1>
        
        <?php if (count($entries) > 0): ?>
            <div class="list-container">
                <?php foreach ($entries as $entry): ?>
                    <div class="list-item">
                        <a target="_blank" rel="noopener noreferrer" href="<?php echo htmlspecialchars($entry['html']); ?>" title="<?php echo htmlspecialchars($entry['name']); ?>">
                            <?php echo htmlspecialchars($entry['name']); ?>
                        </a>
                    </div>
                <?php endforeach; ?>
            </div>
            <div class="info">
                <?php echo count($entries); ?> Einträge gefunden
                <?php if ($latestDate): ?>
                    <br><small>Zuletzt aktualisiert: <?php echo $latestDate->format('d.m.Y \u\m H:i:s'); ?> Uhr</small>
                <?php endif; ?>
            </div>
        <?php else: ?>
            <div class="info">
                ⚠️ Keine HTML-Dateien mit entsprechenden XML-Dateien gefunden
            </div>
        <?php endif; ?>
    </div>
</body>
</html>

<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Результаты анализа протоколов</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        body {
            background-color: #f5f7fa;
            color: #333;
            line-height: 1.6;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }
        
        header {
            background: linear-gradient(135deg, #2c3e50, #4a6491);
            color: white;
            padding: 25px 30px;
        }
        
        h1 {
            font-size: 24px;
            margin-bottom: 10px;
        }
        
        .legend {
            display: flex;
            gap: 20px;
            margin-top: 20px;
            flex-wrap: wrap;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
        }
        
        .color-box {
            width: 16px;
            height: 16px;
            border-radius: 3px;
        }
        
        .red { background-color: #ff5252; }
        .yellow { background-color: #ffd740; }
        .gray { background-color: #b0bec5; }
        
        .content {
            padding: 30px;
        }
        
        .steel-grade {
            margin-bottom: 40px;
        }
        
        h2 {
            font-size: 20px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e0e0e0;
            color: #2c3e50;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        }
        
        th {
            background-color: #f1f5f9;
            padding: 14px 16px;
            text-align: left;
            font-weight: 600;
            color: #2c3e50;
            border-bottom: 1px solid #e0e0e0;
        }
        
        td {
            padding: 12px 16px;
            border-bottom: 1px solid #f0f0f0;
        }
        
        tr:hover {
            background-color: #f9fbfd;
        }
        
        .sample-name {
            font-weight: 500;
        }
        
        .processed-files {
            background-color: #f8f9fa;
            padding: 25px 30px;
            border-top: 1px solid #eaecef;
        }
        
        .file-info {
            display: flex;
            align-items: center;
            gap: 15px;
            margin-top: 15px;
        }
        
        .file-icon {
            width: 40px;
            height: 40px;
            background-color: #e3f2fd;
            border-radius: 5px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #1976d2;
            font-weight: bold;
        }
        
        .file-details {
            flex: 1;
        }
        
        .file-name {
            font-weight: 500;
        }
        
        .file-size {
            font-size: 14px;
            color: #666;
        }
        
        @media (max-width: 768px) {
            .container {
                border-radius: 0;
            }
            
            .content {
                padding: 20px;
            }
            
            table {
                display: block;
                overflow-x: auto;
            }
            
            .legend {
                flex-direction: column;
                gap: 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Результаты анализа протоколов</h1>
            <div class="legend">
                <div class="legend-item">
                    <div class="color-box red"></div>
                    <span>Красный — отклонение от норм</span>
                </div>
                <div class="legend-item">
                    <div class="color-box yellow"></div>
                    <span>Желтый — пограничное значение</span>
                </div>
                <div class="legend-item">
                    <div class="color-box gray"></div>
                    <span>Серый — нормативные требования</span>
                </div>
            </div>
        </header>
        
        <div class="content">
            <div class="steel-grade">
                <h2>Марка стали: 12Х1МФ</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Образец</th>
                            <th>Предел прочности, МПа</th>
                            <th>Предел текучести, МПа</th>
                            <th>Относительное удлинение, %</th>
                            <th>Ударная вязкость, Дж/см²</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="sample-name">НГ ШПП 4</td>
                            <td><span style="color: red;">125 МПа</span></td>
                            <td><span style="color: gray;">250 МПа</span></td>
                            <td><span style="color: yellow;">95%</span></td>
                            <td><span style="color: gray;">22 Дж</span></td>
                        </tr>
                        <tr>
                            <td class="sample-name">НБ ШПП 6</td>
                            <td><span style="color: gray;">275 МПа</span></td>
                            <td><span style="color: gray;">260 МПа</span></td>
                            <td><span style="color: gray;">100%</span></td>
                            <td><span style="color: red;">18 Дж</span></td>
                        </tr>
                        <tr>
                            <td class="sample-name">НА ШПП 4</td>
                            <td><span style="color: gray;">290 МПа</span></td>
                            <td><span style="color: yellow;">275 МПа</span></td>
                            <td><span style="color: gray;">102%</span></td>
                            <td><span style="color: gray;">25 Дж</span></td>
                        </tr>
                        <tr>
                            <td class="sample-name">НА ЦРЯД 57 труба_ПТКМ</td>
                            <td><span style="color: gray;">280 МПа</span></td>
                            <td><span style="color: gray;">265 МПа</span></td>
                            <td><span style="color: gray;">98%</span></td>
                            <td><span style="color: gray;">24 Дж</span></td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div class="steel-grade">
                <h2>Марка стали: 12Х18Н12Т</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Образец</th>
                            <th>Предел прочности, МПа</th>
                            <th>Относительное удлинение, %</th>
                            <th>Ударная вязкость, Дж/см²</th>
                            <th>Стабильность структуры</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="sample-name">НГ 28_КПП ВД</td>
                            <td><span style="color: gray;">520 МПа</span></td>
                            <td><span style="color: gray;">48%</span></td>
                            <td><span style="color: yellow;">110 Дж</span></td>
                            <td><span style="color: gray;">Стабильно</span></td>
                        </tr>
                        <tr>
                            <td class="sample-name">НБ 32_КПП ВД</td>
                            <td><span style="color: red;">480 МПа</span></td>
                            <td><span style="color: gray;">50%</span></td>
                            <td><span style="color: gray;">150 Дж</span></td>
                            <td><span style="color: gray;">Стабильно</span></td>
                        </tr>
                        <tr>
                            <td class="sample-name">НВ 46_КПП ВД</td>
                            <td><span style="color: gray;">510 МПа</span></td>
                            <td><span style="color: yellow;">46%</span></td>
                            <td><span style="color: gray;">140 Дж</span></td>
                            <td><span style="color: red;">Нестабильно</span></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="processed-files">
            <h2>Обработанные файлы</h2>
            <div class="file-info">
                <div class="file-icon">DOCX</div>
                <div class="file-details">
                    <div class="file-name">46. Пшеченкова, Ириклинская ГРЭС.docx</div>
                    <div class="file-size">100.5KB</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>

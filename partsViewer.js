// partsViewer.js

// URLからビルド情報を解析する関数
function parseBuildFromUrl(url) {
    try {
        const urlObj = new URL(url);
        const buildParam = urlObj.searchParams.get('build');
        if (buildParam) {
            return buildParam.split('-').map(Number);
        }
    } catch (e) {
        console.error("URLの解析中にエラーが発生しました:", e);
    }
    return null;
}

// DOCXデータからパーツ情報をパースする関数
function parseDocxPartsData(docxContent) {
    const lines = docxContent.split('\r\n').filter(line => line.trim() !== '');
    // 最初の行に "No." が含まれている場合、それをヘッダーから除外します。
    const headers = lines[0].split(',').map(h => h.trim()).filter(h => h !== 'No.');
    const parts = [];

    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.trim());
        const part = {};
        let valueIndex = 0;
        // "No." 列がある場合はスキップするため、values のインデックスを調整します。
        // これは、values[0]がNo.であると仮定しています。
        if (lines[0].startsWith('No.')) {
            valueIndex = 1;
        }

        headers.forEach((header, index) => {
            if (header === "ENLoad" || header === "Weight") {
                part[header] = parseFloat(values[valueIndex + index]) || undefined; // 数値に変換できない場合は undefined
            } else {
                part[header] = values[valueIndex + index];
            }
        });
        parts.push(part);
    }
    return parts;
}


// DOCXファイルからパーツ情報を取得する関数
async function getPartsDataFromDocx() {
    try {
        const response = await fetch('AC6_partsdate-EN-WEIGHT.docx'); // DOCXファイル名
        if (!response.ok) {
            throw new Error(`HTTPエラー！ステータス: ${response.status}`);
        }
        const docxTextContent = await response.text(); 
        return parseDocxPartsData(docxTextContent);

    } catch (e) {
        console.error("パーツデータの読み込み中にエラーが発生しました:", e);
        return null;
    }
}


// ビルド情報を表示する関数
async function displayBuild() {
    const buildIndices = parseBuildFromUrl(window.location.href);
    const partsData = await getPartsDataFromDocx(); // DOCXからデータを取得する関数を呼び出す
    const buildPartsContainer = document.getElementById('build-parts');
    const totalEnLoadSpan = document.getElementById('total-en-load');
    const totalWeightSpan = document.getElementById('total-weight');
    const errorMessageDiv = document.getElementById('error-message');

    if (!buildIndices || !partsData) {
        errorMessageDiv.classList.remove('hidden');
        return;
    }

    buildPartsContainer.innerHTML = ''; // 以前の内容をクリア
    let totalEnLoad = 0;
    let totalWeight = 0;

    buildIndices.forEach(index => {
        if (index >= 0 && index < partsData.length) {
            const part = partsData[index];
            if (part) {
                const partElement = document.createElement('div');
                partElement.classList.add('part-item');
                partElement.innerHTML = `
                    <span class="part-name">${part.Name} (${part.Kind})</span>
                    <span class="part-stats">
                        EN負荷: ${part.ENLoad !== undefined ? part.ENLoad.toFixed(1) : 'N/A'} |
                        重量: ${part.Weight !== undefined ? part.Weight.toFixed(1) : 'N/A'}
                    </span>
                `;
                buildPartsContainer.appendChild(partElement);

                if (typeof part.ENLoad === 'number' && !isNaN(part.ENLoad)) {
                    totalEnLoad += part.ENLoad;
                }
                if (typeof part.Weight === 'number' && !isNaN(part.Weight)) {
                    totalWeight += part.Weight;
                }
            }
        }
    });

    totalEnLoadSpan.textContent = totalEnLoad.toFixed(1);
    totalWeightSpan.textContent = totalWeight.toFixed(1);

    if (buildPartsContainer.innerHTML === '') {
        errorMessageDiv.classList.remove('hidden');
    } else {
        errorMessageDiv.classList.add('hidden');
    }
}

// ページ読み込み時にビルド情報を表示
window.onload = displayBuild;

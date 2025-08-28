// partsViewer.js

// URLからビルド情報を解析する関数
function parseBuildFromUrl(url) {
    try {
        const urlObj = new URL(url);
        const buildParam = urlObj.searchParams.get('build');
        if (buildParam) {
            const indices = buildParam.split('-').map(part => {
                const num = Number(part);
                if (isNaN(num)) {
                    throw new Error(`不正なビルドID形式です: "${part}"は数値ではありません。`);
                }
                return num;
            });
            return indices;
        }
    } catch (e) {
        // エラーメッセージをコンソールに出力し、詳細なエラーを返す
        console.error("URLの解析中にエラーが発生しました:", e);
        throw new Error(`URL解析エラー: ${e.message || "不明なエラー"}`);
    }
    return null; // buildパラメーターがない場合はnullを返す
}

// DOCXデータからパーツ情報をパースする関数
function parseDocxPartsData(docxContent) {
    const lines = docxContent.split('\r\n').filter(line => line.trim() !== '');
    if (lines.length === 0) {
        throw new Error('DOCXファイルにデータ行がありません。');
    }
    const headers = lines[0].split(',').map(h => h.trim()).filter(h => h !== 'No.');

    if (headers.length === 0) {
        throw new Error('DOCXファイルに有効なヘッダーがありません。');
    }

    const parts = [];
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.trim());
        const part = {};
        let valueIndex = 0;

        if (lines[0].startsWith('No.')) {
            valueIndex = 1; // "No." 列をスキップ
        }

        if ((valueIndex + headers.length) > values.length) {
            // データ行の列数がヘッダーの列数と一致しない場合
            console.warn(`警告: DOCXデータ行 ${i + 1} の形式が不正です。スキップします。`);
            continue; // この行をスキップして次の行に進む
        }

        headers.forEach((header, index) => {
            const val = values[valueIndex + index];
            if (header === "ENLoad" || header === "Weight") {
                part[header] = parseFloat(val);
                if (isNaN(part[header]) && val !== '') { // 空文字列はNaNになるので、空でない場合のみエラーとみなす
                     console.warn(`警告: データ行 ${i + 1} の '${header}' の値 "${val}" は数値としてパースできません。`);
                     part[header] = undefined; // undefinedに設定
                } else if (val === '') {
                    part[header] = undefined; // 明示的に空の数値はundefined
                }
            } else {
                part[header] = val;
            }
        });
        parts.push(part);
    }
    if (parts.length === 0 && lines.length > 1) { // ヘッダー以外にデータ行があったにも関わらずパースされたパーツがない場合
        throw new Error('DOCXファイルから有効なパーツデータを抽出できませんでした。データ形式を確認してください。');
    }
    return parts;
}


// DOCXファイルからパーツ情報を取得する関数
async function getPartsDataFromDocx() {
    try {
        const response = await fetch('AC6_partsdate-EN-WEIGHT.docx'); // DOCXファイル名
        if (!response.ok) {
            // HTTPエラーの詳細メッセージ
            throw new Error(`パーツファイル読み込みエラー: HTTPステータス ${response.status} - ${response.statusText}`);
        }
        const docxTextContent = await response.text();
        return parseDocxPartsData(docxTextContent);

    } catch (e) {
        // fetchやparseDocxPartsDataからのエラーを補足し、より詳細なメッセージを追加
        console.error("パーツデータの取得中にエラーが発生しました:", e);
        throw new Error(`パーツデータ取得エラー: ${e.message || "不明なエラー"}`);
    }
}


// ビルド情報を表示する関数
async function displayBuild(url) {
    const buildPartsContainer = document.getElementById('build-parts');
    const totalEnLoadSpan = document.getElementById('total-en-load');
    const totalWeightSpan = document.getElementById('total-weight');
    const errorMessageDiv = document.getElementById('error-message');
    const loadingIndicator = document.getElementById('loading-indicator');

    // UIをリセットし、ローディングインジケーターを非表示にする（初期状態または無効なURLの場合）
    buildPartsContainer.innerHTML = '';
    totalEnLoadSpan.textContent = '0.0';
    totalWeightSpan.textContent = '0.0';
    errorMessageDiv.classList.add('hidden');
    loadingIndicator.classList.add('hidden'); // まずローディングインジケーターを非表示にする

    // URLが提供されていない場合、何も表示せずに終了する (初期ロード時のクリーンな状態)
    if (!url) {
        return;
    }

    loadingIndicator.classList.remove('hidden'); // URLが提供されたのでローディングインジケーターを表示

    let buildIndices = null;
    try {
        buildIndices = parseBuildFromUrl(url);
    } catch (e) {
        errorMessageDiv.textContent = `URL解析エラー: ${e.message}`;
        errorMessageDiv.classList.remove('hidden');
        loadingIndicator.classList.add('hidden');
        return;
    }

    // buildパラメーター自体がない場合
    if (!buildIndices) {
        errorMessageDiv.textContent = 'URLに有効なビルド情報が見つかりませんでした。"?build=..."形式のURLを入力してください。';
        errorMessageDiv.classList.remove('hidden');
        loadingIndicator.classList.add('hidden');
        return;
    }

    let partsData = null;
    try {
        partsData = await getPartsDataFromDocx();
    } catch (e) {
        errorMessageDiv.textContent = `パーツデータの読み込みエラー: ${e.message}`;
        errorMessageDiv.classList.remove('hidden');
        loadingIndicator.classList.add('hidden');
        return;
    }

    if (!partsData || partsData.length === 0) {
        errorMessageDiv.textContent = 'パーツデータが空か、正しく読み込めませんでした。`AC6_partsdate-EN-WEIGHT.docx`ファイルの内容を確認してください。';
        errorMessageDiv.classList.remove('hidden');
        loadingIndicator.classList.add('hidden');
        return;
    }

    let totalEnLoad = 0;
    let totalWeight = 0;
    let partsFound = false;

    buildIndices.forEach(index => {
        if (index >= 0 && index < partsData.length) {
            const part = partsData[index];
            if (part) {
                partsFound = true;
                const partElement = document.createElement('div');
                partElement.classList.add('part-item');
                partElement.innerHTML = `
                    <span class="part-name">${part.Name || '不明なパーツ'} (${part.Kind || '不明'})</span>
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
        } else {
            // 範囲外のインデックスに対する警告
            console.warn(`警告: ビルドID ${index} はパーツデータ範囲外です。`);
        }
    });

    totalEnLoadSpan.textContent = totalEnLoad.toFixed(1);
    totalWeightSpan.textContent = totalWeight.toFixed(1);

    if (!partsFound) {
        errorMessageDiv.textContent = '指定されたビルドIDに該当するパーツが見つかりませんでした。ビルドIDを確認してください。';
        errorMessageDiv.classList.remove('hidden');
    } else {
        errorMessageDiv.classList.add('hidden'); // 成功した場合はエラーメッセージを隠す
    }

    loadingIndicator.classList.add('hidden'); // すべての処理が完了したら非表示にする
}

// ページ読み込み時の初期化
window.onload = function() {
    const loadButton = document.getElementById('load-build-button');
    const urlInput = document.getElementById('build-url');
    const errorMessageDiv = document.getElementById('error-message');

    // 読み込みボタンがクリックされたときの処理
    loadButton.addEventListener('click', () => {
        const inputUrl = urlInput.value.trim();
        if (inputUrl) {
            displayBuild(inputUrl);
        } else {
            errorMessageDiv.textContent = 'URLを入力してください。';
            errorMessageDiv.classList.remove('hidden');
            document.getElementById('loading-indicator').classList.add('hidden');
            // 他の表示をクリア
            document.getElementById('build-parts').innerHTML = '';
            document.getElementById('total-en-load').textContent = '0.0';
            document.getElementById('total-weight').textContent = '0.0';
        }
    });

    // ページロード時にURLにビルドパラメータがある場合は自動で読み込む
    if (window.location.search.includes('build=')) {
        urlInput.value = window.location.href;
        displayBuild(window.location.href);
    } else {
        // 'build'パラメーターがない場合、エラーメッセージを表示せず、ローディングインジケーターも表示しない
        displayBuild(null);
    }
};

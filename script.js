function setScript(srcList) {
    return new Promise((resolve, reject) => {
        if (srcList.length == 0) {
            reject();
            return;
        }
        const src = srcList.shift();
        if (!src) {
            reject();
            return;
        }
        const srcEl = document.createElement('script');
        srcEl.onload = () => {
            resolve()
        }
        srcEl.onerror = async () => {
            await setScript(srcList)
        }
        srcEl.src = src;
        document.head.appendChild(srcEl);
    })

}
function getCoinData(coinSelector) {
    const AddressesRicherThan = ['Addresses richer than 1 USD', 'Addresses richer than 100 USD', 'Addresses richer than 1,000 USD', 'Addresses richer than 10,000 USD'];
    const WealthDistributionTop = ['Wealth Distribution Top 10 addesses', 'Wealth Distribution Top 100 addesses', 'Wealth Distribution Top 1,000 addesses', 'Wealth Distribution Top 10,000 addesses'];

    const data = {
        'Date/Time': dayjs().format('YYYY/M/D H:m'),
    };
    document.querySelectorAll(coinSelector).forEach(item => {
        const title = item.parentElement.querySelector('td').innerText;
        if (title.indexOf('1/100/1,000/10,000 USD') > -1) {
            item.innerText.split(' / ').forEach((usd, index) => {
                data[AddressesRicherThan[index]] = usd
            })
        } else if (title.indexOf('10/100/1,000/10,000 addesses') > -1) {
            item.innerText.replace(' Total', '').split(' / ').forEach((add, index) => {
                data[WealthDistributionTop[index]] = add
            })
        } else if (title.indexOf('Price') > -1) {
            data[title] = item.innerText.split('\n')[0]
        } else if (title && title.length > 0) {
            data[title] = item.innerText.replaceAll('\n', ' ');
        }
    })
    console.log(data);
    return data;
}
function getExcel(coinDataList) {
    const workbook = new ExcelJS.Workbook();

    coinDataList.forEach(coinData => {
        const sheet = workbook.addWorksheet(coinData.name);

        sheet.addTable({
            name: coinData.name,
            ref: 'A1',
            columns: [{ name: 'key' }, { name: 'value' }],
            rows: []
        });

        Object.entries(coinData.data).forEach(item => {
            if (item[0].length > 0 && item[1].length > 0) {
                sheet.addRow(item)
            }
        })
    });


    workbook.xlsx.writeBuffer().then((content) => {
        const link = document.createElement("a");
        const blobData = new Blob([content], {
            type: "application/vnd.ms-excel;charset=utf-8;"
        });
        link.download = `${dayjs().format('YYYY/M/D H:m')}.xlsx`;
        link.href = URL.createObjectURL(blobData);
        link.click();
    });
}
async function init() {

    if (window.location.host !== 'bitinfocharts.com') {
        alert('此程式僅限用於 bitinfocharts');
        window.location.href = "https://bitinfocharts.com/"
        return;
    }

    await setScript(['https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js']);
    await setScript(['https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js']);


    const btcData = getCoinData('.coin.c_btc');
    const dogeData = getCoinData('.coin.c_doge');

    getExcel([
        {
            name: 'BTC',
            data: btcData,
        },
        {
            name: 'Dogecoin',
            data: dogeData,
        },
    ])

}
init();
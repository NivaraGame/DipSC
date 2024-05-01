let dict = {
    source: '',
    template: '',
    dist: '',
    sheet: '',
    group: 0
}

let sheets = []

async function choose_file(type) {
    dict[type] = await eel.choose_file()()
    document.getElementById(type).innerText = dict[type]
    if (type === 'source') {
        sheets = await eel.get_sheets(dict['source'])()
        get_radio()
    }
}

function get_radio() {
    document.getElementById('sheet_name').innerHTML = ''
    sheets.forEach(item => {
        document.getElementById('sheet_name').innerHTML += `
        <label class="pl-5">
            <input type="radio" name="some" onclick='choose_sheet("${item}")'/>
            ${item}
        </label>`
    })
}

async function choose_folder() {
    dict['dist'] = await eel.choose_folder()()
    document.getElementById('dist').innerText = dict['dist'] + '/'
}

function choose_sheet(sheet) {
    dict['sheet'] = sheet
}

function set_group() {
    dict['group'] = document.getElementById('group').value
}

async function start() {
    document.getElementById('add').innerText = 'Зачекайте'
    await eel.start(dict)()
    document.getElementById('add').innerText = 'Виконано'
}

function callBack() {
    console.log(dict)
}
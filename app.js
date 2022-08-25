const form = document.querySelector('form');
const cardTypeSelect = document.querySelector('#card-type');
const genderFieldset = document.querySelector('.gender-fieldset');
const yearsFieldset = document.querySelector('.years-fieldset');

form.addEventListener('submit', onSubmit);
cardTypeSelect.addEventListener('input', onTypeSelect);

function onTypeSelect({ target }) {
    const cardType = target.value;
    if (cardType === 'anniversary') {
        genderFieldset.style.display = 'none';
        yearsFieldset.style.display = 'initial';
        return;
    }
    genderFieldset.style.display = 'initial';
    yearsFieldset.style.display = 'none';
}

function onSubmit(e) {
    e.preventDefault();
    const data = new FormData(form);
    const name = data.get('name').trim();

    if (cardTypeSelect.value === 'anniversary') {
        generateAnniversaryCard(name, data.get('years'));
        return;
    }

    generateBirthdayCard(name, data.get('gender'));
}

function generateBirthdayCard(name, gender) {
    const outputFileName = `Happy Birthday ${capitalize(name)}.pptx`;
    const blankURL = `./blanks/birthday/${gender}_blank.pptx`;
    const template = {
        name: name.toUpperCase()
    }

    generate(blankURL, outputFileName, template);
}

function generateAnniversaryCard(name, years) {
    const outputFileName = `Happy anniversary ${capitalize(name)}.pptx`;
    const blankURL = `./blanks/anniversary/anniversary_${years}.pptx`;
    const template = {
        name: capitalize(name)
    };

    generate(blankURL, outputFileName, template);
}

function generate(blankURL, outputFileName, template) {
    loadFile(
        blankURL,
        function (error, content) {
            if (error) {
                throw error;
            }
            let zip = new PizZip(content);
            let doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render(template);

            let out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                // compression: "DEFLATE",
            });
            saveAs(
                out,
                outputFileName
            );
        }
    );
}

function loadFile(url, callback) {
    PizZipUtils.getBinaryContent(url, callback);
}

function capitalize(str) {
    return str[0].toUpperCase() + str.slice(1).toLowerCase()
}

const form = document.querySelector('form');
form.addEventListener('submit', onSubmit)

function onSubmit(e) {
    e.preventDefault();
    const data = new FormData(form);

    generate(data.get('name'), data.get('gender'));
}

function loadFile(url, callback) {
    PizZipUtils.getBinaryContent(url, callback);
}
function generate(name, gender) {
    const outputFileName = `Happy Birthday ${capitalize(name)}.pptx`;
    const blankURL = `./${gender}_blank.pptx`;

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

            // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
            doc.render({
                name: name.toUpperCase()
            });

            let out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                compression: "DEFLATE",
            });
            // Output the document using Data-URI
            saveAs(
                out,
                outputFileName
            );
        }
    );
}

function capitalize(str) {
    return str[0].toUpperCase() + str.slice(1).toLowerCase()
}

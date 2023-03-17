// Load all prerequisites
const { readFile, writeFile, rename } = require('fs').promises;
const filesystem = require('fs');
const translate = require('google-translate-api-x');
const http = require('http');
const formidable = require('formidable');
//const officeParser = require('officeparser');
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const port = 8080;
const sep = "\n\@\n";

// Function to send 
function send_index(index_html, res) {
    res.writeHead(200, {'Content-Type': 'text/html' });
    res.write(index_html);
    res.end();
}

async function translate_file(filepath) {
    text_id = 0
    // const filepath = files.multipleFiles.filepath;
    console.log("Converting pptx to JSON")
    working_json = await pptx2json.toJson(filepath);
    text_to_translate = []
    extracted_spaces = []
    // A bit wasteful, but we can leverage the recursive function within JSON.stringify to find the text within each page, as well as putting it back.
    // First, grab all of the text within the "a:t" tags
    console.log("First stringification...")
    JSON.stringify(working_json, (key, nested_obj) => {
        if (nested_obj && nested_obj["a:t"] != undefined) {
            txt = nested_obj["a:t"][0];
            if (txt.length > 0) {
                nested_obj.text_id = text_id;
                text_to_translate.push(txt);
                text_id++;

                // Save off the number of pre and post spaces
                // because Google Translate strips them.
                stripped_text = txt.trim();
                pre_spaces = txt.indexOf(stripped_text);
                post_spaces = txt.length - (stripped_text.length + pre_spaces);
                extracted_spaces.push([pre_spaces, post_spaces]);
            }
        }
        return nested_obj;
    });

    // joined_text = text_to_translate.join(sep);
    console.log({"Extracted text": text_to_translate});

    // Then, translate all the text in one bulk translate
    console.log("Sending to Google for translations...")
    const res = await translate(text_to_translate, { from: 'en', to: 'es' });
    translated_text = []
    res.forEach((val, idx) => {
        translated_text.push(val.text)
    })

    text_id = undefined;

    console.log("Second stringification...");
    // Finally, insert the translated text back into 
    JSON.stringify(working_json, (key, nested_obj) => {
        if (nested_obj && nested_obj["a:t"] != undefined && nested_obj.text_id != undefined) {
            if (text_id == undefined) {
                text_id = nested_obj.text_id;
            }
            translated_string = translated_text[nested_obj.text_id];
            pre_spaces = ' '.repeat(extracted_spaces[nested_obj.text_id][0]);
            post_spaces = ' '.repeat(extracted_spaces[nested_obj.text_id][1]);
            text_to_insert = pre_spaces + translated_string + post_spaces;
            nested_obj["a:t"] = [text_to_insert];
        }
        return nested_obj;
    });

    console.log("Converting back into a pptx file, overwriting the previous file");
    // Currently overwrite the uploaded file. I'm sure this won't have any negative consequences /s
    await pptx2json.toPPTX(working_json, {'file': filepath});
}

function accept_form_data(req, res) {
    if (req.method.toLowerCase() !== "post") {
        console.log(req);
        res.writeHead(400, {'Content-Type': 'text/html'});
        res.write("Bad request. This endpoint is not supported for non-POST requests");
    } else {
        // Parse and save any form data
        const form = formidable({ multiples: true });

        console.log("File received, attempting to parse...")
        form.parse(req, (err, fields, files) => {
            if (err) {
                res.writeHead(err.httpCode || 400, {'Content-Type': 'text/plain' });
                res.end(String(err));
                return;
            }
            if (files.multipleFiles.length == undefined) {
                // Dealing with a single file
                console.log("Working with a single file, beginning file translation process...")
                const filepath = files.multipleFiles.filepath;
                
                translate_file(filepath).then(() => {
                    console.log("Translation complete, sending the data back to the user!");
                    stat = filesystem.statSync(filepath);
                    res.writeHead(200, { 
                        'Content-Length': stat.size, 
                        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        'Content-Disposition': 'attachment; filename="' + files.multipleFiles.originalFilename + '"'
                    });
                    readStream = filesystem.createReadStream(filepath)
                    readStream.pipe(res);
                });
            }
            else {
                // Dealing with multiple files
                // for (var i = 0; i < files.multipleFiles.length; i++) {

                // }
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ fields, files}, null, 2));
            }
        });
    }
}

async function initialize_server() {
    const index_html = await readFile('./index.html', 'utf-8');
    const server = http.createServer((req, res) => {
        console.log(req.url, req.method);
        switch (req.url) {
            case '/api/submit_file':
                accept_form_data(req, res);
                break;
            case '/':
            default:
                send_index(index_html, res);
        }
    });

    server.listen(port);
}

initialize_server();

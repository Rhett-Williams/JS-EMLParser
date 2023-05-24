const EmlParser = require('eml-parser');
const fs = require('fs');
const fsp = require('fs').promises;
const { JSDOM } = require('jsdom');
const path = require('path');
const {convert} = require('html-to-text');
const readlineSync = require('readline-sync');

const defaultPath = `${path.dirname(process.execPath)}/${readlineSync.question('Enter your file name: ')}`;


function addTextAsHtmlToEml(emlFilePath) {
  // Read the original EML file
  fs.readFile(emlFilePath, 'utf8', (err, data) => {
    if (err) {
      console.error('Error reading the EML file:', err);
      return;
    }

    // Find the HTML content within the EML file
    const htmlStart = data.indexOf('<html');
    const htmlEnd = data.lastIndexOf('</html>');
    if (htmlStart === -1 || htmlEnd === -1) {
      console.error('No HTML content found in the EML file.');
      return;
    }

    const htmlContent = data.slice(htmlStart, htmlEnd + 7); // Include the closing tag

    // Convert the HTML to plain text
    const textContent = convert(htmlContent, {
      wordwrap: false,
      ignoreHref: true,
      ignoreImage: true,
    });

    // Modify the EML file to include "textAsHtml" content
    const modifiedEmlContent = data.replace(
      'Content-Type: text/html;',
      'Content-Type: multipart/alternative;\r\n\tboundary="==123=="\r\n\r\n--==123==\r\nContent-Type: text/plain;\r\n\tcharset="utf-8"\r\nContent-Transfer-Encoding: 7bit\r\n\r\n' + textContent + '\r\n\r\n--==123==\r\nContent-Type: text/html;\r\n\tcharset="utf-8"\r\nContent-Transfer-Encoding: 7bit\r\n\r\n' + htmlContent + '\r\n\r\n--==123==--'
    );

    // Save the modified EML file
    const modifiedEmlFilePath = emlFilePath.replace('.eml', '_modified.eml');
    fs.writeFile(modifiedEmlFilePath, modifiedEmlContent, 'utf8', (err) => {
      if (err) {
        console.error('Error saving the modified EML file:', err);
        return;
      }
    });
    runScript(modifiedEmlFilePath)
  });
}

async function getCss(textContent, callback, index, htmlIndex) {
  try {

    return new Promise((resolve, reject) => {
      new EmlParser(fs.createReadStream(defaultPath))
        .parseEml()
        .then(result => {
          const { document } = new JSDOM(result.html).window;
          const elements = document.querySelectorAll('*');

          const normalizedTextContent = normalizeTextContent(textContent);
          if ((normalizedTextContent === '') || (normalizedTextContent === ' ') || (!normalizedTextContent)){
            console.log('Element not found.', index, normalizedTextContent);
            callback(undefined, htmlIndex);
            return
          }
          let foundElement = null;
          let foundIndex = 0;
          for (let i = htmlIndex; i < elements.length; i++) {
            const element = elements[i];
            const elementTextContent = normalizeTextContent(element.textContent);
            if (elementTextContent === normalizedTextContent) {
              foundElement = element;
              foundIndex = i;
              break;
            }
          }

          if (foundElement) {

            const ancestorElements = getAncestorElements(foundElement, 7);

            const anc = ancestorElements[6]
            const innerHTML = anc?.innerHTML;
            callback(innerHTML, foundIndex);
          } else {
            console.log('Element not found.', index, normalizedTextContent)
            callback(undefined, htmlIndex);
          }
        })})
  } catch (error) {
    console.log(error);
  }
}

function normalizeTextContent(textContent) {
  return textContent
    .replace(/&lt;.*?&gt;/g, '')
    .replace(/&quot;/g, "'")
    .replace(/<br\/?>/g, '')
    .replace(/= /g, '')
    .replace(/=/g, '')
    .replaceAll(/&reg;/g, "®")
    .replaceAll(/&ldquo;/g, '“')
    .replace(/&apos;/g, "'")
    .replace(/\n/g, '')
    .replaceAll(' ', '')
    .replace(/\[.*?\]/g, '')
    .replaceAll(/E28093/g, '-')
    .replace(/<[^>]+>/g, '')
    .toLowerCase()
    .normalize('NFKC')
    .trim();
}

function getAncestorElements(element, levels) {
  const ancestors = [];
  let currentElement = element;

  while (currentElement && levels > 0) {
    ancestors.push(currentElement);
    currentElement = currentElement.parentNode;
    levels--;
  }

  return ancestors;
}

function getTextContentArray(filePath) {
  return new Promise((resolve, reject) => {
    new EmlParser(fs.createReadStream(filePath))
      .parseEml()
      .then(result => {
        const subjectDirectoryName = result.subject.replaceAll(" ", "_");

        for (const attachment of result.attachments) {
          if (attachment.contentType.startsWith('image/')) {
            // Generate a unique file name for the image
            const extension = path.extname(attachment.filename);
            const imageFileName = `${Date.now()}-${Math.floor(Math.random() * 10000)}${extension}`;
            // const filePath = path.join(process.cwd(), imageFileName);
            const folderPath = path.join(path.dirname(process.execPath), `media_${subjectDirectoryName}`);
            const filePath = path.join(folderPath, imageFileName);
            if (!fs.existsSync(folderPath)) {
              fs.mkdirSync(folderPath);
            }
            fs.writeFileSync(filePath, attachment.content);
        
            console.log(`Image saved: ${imageFileName}`);
          }
        }

        if (!result.textAsHtml){
          addTextAsHtmlToEml(filePath)
          return
        }

        var paragraphs = result.textAsHtml.split(/<\/?p>/).filter(function(paragraph) {
          return paragraph.trim().length > 0;
        });
        let htmlData = [];

        function processParagraphs(para, index, htmlIndex) {
          if (index >= para.length) {
            resolve({htmlData, subjectDirectoryName});
            return;
          }

          getCss(para[index], (data, htmlIndex) => {
            htmlData = htmlData.concat(data);
            processParagraphs(para, index + 1, htmlIndex);
          }, index, htmlIndex);
        }

        processParagraphs(paragraphs, 0, 0);
      })
      .catch(err => {
        reject(err);
      });
  });
}

function removeImgTags(text) {
  const regex = /<img\b[^>]*>/gi;
  if(!text)return
  return text.replaceAll(regex, '');
}

function fixHtml(html) {
  var result = html.replace(/<!--[\s\S]*?-->/g, '').replace(/\s?pardot-region="[^"]*"/g, '').replaceAll("{{", "{{{").replaceAll("}}", "}}}").replace(/<o:p>(<\/o:p>)?/g, '');
  return result;
}

function removeDuplicates(strings) {
  const uniqueStrings = new Set();
  const result = [];
  for (let i = 0; i < strings.length; i++) {
    const string = strings[i];
    if (!uniqueStrings.has(string)) {
      uniqueStrings.add(string);
      const removedTags = removeImgTags(string)
      if (removedTags){
        result.push(removedTags);
      }
    }
  }
  return result;
}

function runScript(file){
  getTextContentArray(file)
  .then(data => {
    const res = removeDuplicates(data.htmlData)
    const fixedData = res.map((string) => {
      return fixHtml(string)
    })
    console.log("asd", data.subjectDirectoryName)

    fixedData.map((item, index) => {
      if (!item) return;
      // Generate the file name
      const fileName = `output${index}.html`;
      // Create the file path by joining the current working directory with the file name

      const folderPath = path.join(path.dirname(process.execPath), `outputs_${data.subjectDirectoryName}`);
      const filePath = path.join(folderPath, fileName);
      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath);
      }
      fs.writeFile(filePath, item, (err) => {
        if (err) {
          console.error(err);
        } else {
          console.log(`File saved: ${filePath}`);
        }
      });
    });


  })
  .catch(err => {
    console.error(err);
  });
}

runScript(defaultPath)

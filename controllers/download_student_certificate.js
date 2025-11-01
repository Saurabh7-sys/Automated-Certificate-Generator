/*
Setup:

npm install pizzip --save (used to zip the buffer --new PizZip(buffer))
npm install docxtemplater --save (for substituting the value)
npm install docx-merger --save (Used to merge the document)
npm install xml2js --save (converting xml to js word => js)
npm install html-to-text --save (converting html to text html => word)
npm install office-to-pdf --save (converting word to pdf)
npm install libreoffice-convert --save
npm install docxtemplater-image-module-free --save (for substituting with images)
*/

const controller = {};

const fs = require('fs').promises;
const fss = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const DocxMerger = require('docx-merger');
const xml2js = require('xml2js');
const { htmlToText } = require('html-to-text');
const toPdf = require("office-to-pdf");
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
const { promisify } = require('util');
const parseStringPromise = promisify(xml2js.parseString);
const ImageModule = require("docxtemplater-image-module-free");

controller.tempWordFiles = [];

controller.readJSON = async function(file_path) {
	const data = await fs.readFile(file_path, "utf8");
	if (data) {
		return JSON.parse(data);
	}
	return;
};

controller.get_attribute = function(student_info, attribute) {
	let value = '';
	if (student_info.hasOwnProperty(attribute)) {
		value = student_info[attribute];
	}
	return controller.get_value(value, attribute);
};

controller.get_value = function(value, red_text) {
	if (value != null && typeof value != 'undefined') {
		value = value.toUpperCase();
	}
	if (value != '') {
		return value;
	}
	return "<red>" + red_text + "</red>";
};

// Function to replace <red> tags with red-colored text
function replaceRedTags(docx) {
	const xml = docx.getZip().files['word/document.xml'].asText();

	// console.log("entering replaceRedTags");

	// Replace &lt;red&gt; and &lt;/red&gt; with XML markup for red-colored and bold text
	const updatedXml = xml.replace(/&lt;red&gt;(.*?)&lt;\/red&gt;/g, (match, p1) => {
		// console.log(p1);
		/*
			<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FF0000"/></w:rPr><w:t>hello</w:t></w:r>
		*/
        // return `<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FF0000"/></w:rPr><w:t>${p1}</w:t></w:r>`;
        return `----`;
	});

	// console.log("leaving  replaceRedTags");

	docx.getZip().file('word/document.xml', updatedXml);
};

// Function to replace <red> tags with red-colored text and bold formatting
function replaceRedTags2(docx) {
    const xml = docx.getZip().files['word/document.xml'].asText();

    // Replace <red> and </red> with proper XML markup for red-colored and bold text
    const updatedXml = xml.replace(/<red>(.*?)<\/red>/g, (match, p1) => {
        return `
            <w:r>
                <w:rPr>
                    <w:color w:val="FF0000"/>
                    <w:b/>
                </w:rPr>
                <w:t>${p1}</w:t>
            </w:r>`;
    });

    // Update the document.xml with the new content
    docx.getZip().file('word/document.xml', updatedXml);
}

controller.run = async function(req, res) {

	const input_file_path = req.body.input_file_path;
	const output_file_path = req.body.output_file_path;
	const data_directory = req.body.data_directory + '/';
	const output_file_path_pdf = output_file_path.replace('.docx', '.pdf');

	controller.tempWordFiles = [];

	let all_json_files = await fs.readdir(data_directory, (err) => {
		console.log("error ", err);
	});

	// console.log(all_json_files);

	let student_info;

	for (let i = 0; i < all_json_files.length; i++) {

		student_info = await controller.readJSON(data_directory + all_json_files[i]);

		if (!student_info) {
			console.log(all_json_files[i], " data is not present");
			// return; // CHANGE: THROW ERROR
		}

		student_info.DOB = student_info.DOB.replace(/\//g, "-");
		student_info.DOJ = student_info.DOJ.replace(/\//g, "-");
		student_info.DOL = student_info.DOL.replace(/\//g, "-");
		student_info.DATE_OF_APPLICATION_FOR_TC = student_info.DATE_OF_APPLICATION_FOR_TC.replace(/\//g, "-");
		student_info.TC_ISSUE_DATE = student_info.TC_ISSUE_DATE.replace(/\//g, "-");
		student_info.DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER = student_info.DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER.replace(/\//g, "-");

		// Creating all temp documents seperately
		await controller.generateDocuments(data_directory, student_info, input_file_path, i);
	}

	// Merging the all temp documents in one

	await controller.merging_doc(data_directory, input_file_path, output_file_path, false);

	// Deleting all temp documents
	controller.tempWordFiles.forEach(async file => await fs.unlink(file));

	// Converting Docx to Pdf
	await controller.convertDocxToPdf_officeTopdf(output_file_path, output_file_path_pdf);

	let download_file = output_file_path.replace('/var/www/html/sites/repos/dspl-live/server/', '');

	res.json({
		response_code: 1,
		response_message: 'OK',
		download_file: download_file
	});
};



// Generating temp Docx documents
controller.generateDocuments = async function(data_directory, student_info, input_file_path, temp_file_number) {

	try {

	// Read the binary content of the input Word document
	const content = await fs.readFile(input_file_path, 'binary');
	// Create a PizZip instance with the binary content
	const zip = new PizZip(content);

  // 3. Initialize ImageModule
  const imageModule = new ImageModule({
      getImage: (tagValue, tagName) => {
      // Ensure that we get the image file buffer
        let imageBuffer = fss.readFileSync(tagValue);
      return imageBuffer;
    },
      getSize: (imgBuffer, tagValue, tagName) => {
		return data.IMAGE_SIZES[tagName]; // Image size for example
	  }
  });

  // 4. Create a Docxtemplater instance with the zip content and image module
  const doc = new Docxtemplater(zip, {
		modules: [imageModule],
		paragraphLoop: true,
		linebreaks: true,
	});

	student_info.GENDER = student_info.GENDER.trim().toUpperCase();

	let today = new Date();

	let pic_file_path = student_info.IMAGE_FILE_PATH;
	if (pic_file_path != '') {
		pic_file_path = '/var/www/html/sites/erp.decagonsoftware.com/erp-code/server/' + pic_file_path;
	}
	console.log("pic_file_path = ", pic_file_path);

// Using data because in new update setData is depreciate
	let data = ({
		ADMN_NUM: controller.get_attribute(student_info, 'ADMN_NUM'),
		IMAGE_FILE_PATH: pic_file_path,
		// STU_IMAGE: "abc.png",
		// IMAGE_13: "123.png",
		IMAGE_SIZES : {
			IMAGE_FILE_PATH : [130, 164],
			// IMAGE_13 : [20, 80],
		},
		// STU_IMAGE_SIZE: controller.get_attribute(student_info, 'STU_IMAGE_SIZE'),
		STU_NAME: controller.get_attribute(student_info, 'STU_NAME'),
		CLASS_NAME: controller.get_attribute(student_info, 'CLASS_NAME'),
		NATIONALITY: controller.get_attribute(student_info, 'NATIONALITY'),
		BELONGS_TO_ST_SC: controller.get_attribute(student_info, 'BELONGS_TO_ST_SC'),
		LAST_EXAM_WITH_RESULT: controller.get_attribute(student_info, 'LAST_EXAM_WITH_RESULT'),
		FAILED_ONCE_OR_TWICE_IN_SAME_CLASS: controller.get_attribute(student_info, 'FAILED_ONCE_OR_TWICE_IN_SAME_CLASS'),
		SUBJECTS_STUDIED: controller.get_attribute(student_info, 'SUBJECTS_STUDIED'),
		PROMOTED_FOR_TC: controller.get_attribute(student_info, 'PROMOTED_FOR_TC'),
		FEES_PAID_UPTO: controller.get_attribute(student_info, 'FEES_PAID_UPTO'),
		FEE_CONCESSION_DETAILS: controller.get_attribute(student_info, 'FEE_CONCESSION_DETAILS'),
		TOTAL_WORKING_DAYS: controller.get_attribute(student_info, 'TOTAL_WORKING_DAYS'),
		TOTAL_PRESENT_DAYS: controller.get_attribute(student_info, 'TOTAL_PRESENT_DAYS'),
		NCC_CADET_DETAILS: controller.get_attribute(student_info, 'NCC_CADET_DETAILS'),
		GAMES_OR_EXTRACURRICULAR_ACT: controller.get_attribute(student_info, 'GAMES_OR_EXTRACURRICULAR_ACT'),
		ACHIEVEMENT_IN_GAMES_OR_EXTRACURRICULAR_ACT: controller.get_attribute(student_info, 'ACHIEVEMENT_IN_GAMES_OR_EXTRACURRICULAR_ACT'),
		CHARACTER_FOR_TC: controller.get_attribute(student_info, 'CHARACTER_FOR_TC'),
		DATE_OF_APPLICATION_FOR_TC_FORMATTED: controller.get_attribute(student_info, 'DATE_OF_APPLICATION_FOR_TC_FORMATTED'),
		DATE_OF_APPLICATION_FOR_TC: controller.get_attribute(student_info, 'DATE_OF_APPLICATION_FOR_TC'),
		TC_ISSUE_DATE: controller.get_attribute(student_info, 'TC_ISSUE_DATE'),
		REASON_FOR_LEAVING: controller.get_attribute(student_info, 'REASON_FOR_LEAVING'),
		ANY_OTHER_REMARKS: controller.get_attribute(student_info, 'ANY_OTHER_REMARKS'),
		CURRENT_ADDRESS: controller.get_attribute(student_info, 'CURRENT_ADDRESS'),
		PERMANENT_ADDRESS: controller.get_attribute(student_info, 'PERMANENT_ADDRESS'),
		NOW_FORMATTED: controller.get_attribute(student_info, 'NOW_FORMATTED'),

		DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER_MMM_YYYY: controller.get_attribute(student_info, 'DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER_MMM_YYYY'),
		DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER_DD_MMM_YYYY: controller.get_attribute(student_info, 'DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER_DD_MMM_YYYY'),
		TC_ISSUE_DATE_DMY: controller.get_attribute(student_info, 'TC_ISSUE_DATE_DMY'),
		MOTHER_EMAIL: controller.get_attribute(student_info, 'MOTHER_EMAIL'),
		FATHER_EMAIL: controller.get_attribute(student_info, 'FATHER_EMAIL'),
		MOTHER_MOB: controller.get_attribute(student_info, 'MOTHER_MOB'),
		FATHER_MOB: controller.get_attribute(student_info, 'FATHER_MOB'),
		FAMILY_ANNUAL_INCOME: controller.get_attribute(student_info, 'FAMILY_ANNUAL_INCOME'),

		DOJ: controller.get_value(student_info.DOJ, 'Date of Joining'),
		DOL: controller.get_value(student_info.DOL, 'Date of Leaving'),
		DOL_WITH_MONTH_NAME: controller.get_value(student_info.DOL_WITH_MONTH_NAME, "Date of Leaving"),
		DOB: controller.get_value(student_info.DOB, 'Date of Birth'),
		DOB_WORDS: controller.get_value(student_info.DOB_WORDS, 'Date of Birth in words'),
		DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER: controller.get_value(student_info.DOJ_OR_SESSION_START_DATE_WHICHEVER_IS_LATER, 'DOJ or Session Start Date'),
		ISSUED_TC_NUM: controller.get_value(student_info.ISSUED_TC_NUM, 'Issued TC Number'),
		FATHER: controller.get_value(student_info.FATHER, "Father's Name"),
		MOTHER: controller.get_value(student_info.MOTHER, "Mother's Name"),
		SELECTED_SESSION: controller.get_attribute(student_info, 'SELECTED_SESSION'),
		CURRENT_SESSION: controller.get_attribute(student_info, 'CURRENT_SESSION'),
		ROLL_NUMBER: controller.get_attribute(student_info, 'ROLL_NUMBER'),
		BOARD_REGISTRATION_NUMBER: controller.get_attribute(student_info, 'BOARD_REGISTRATION_NUMBER'),
		TOTAL_FEE_PAID: controller.get_attribute(student_info, 'TOTAL_FEE_PAID'),
		SELECTED_FEE_HEAD_NAMES: controller.get_attribute(student_info, 'SELECTED_FEE_HEAD_NAMES'),
		SCHOOL_1: controller.get_attribute(student_info, 'SCHOOL_1'),
		CLASSES_COMPLETED_1: controller.get_attribute(student_info, 'CLASSES_COMPLETED_1'),
		STU_REF_NUMBER: controller.get_attribute(student_info, 'STU_REF_NUMBER'),
		PREVIOUS_SCHOOL: controller.get_attribute(student_info, 'PREVIOUS_SCHOOL'),
		PREVIOUS_SCHOOL_ADDRESS: controller.get_attribute(student_info, 'PREVIOUS_SCHOOL_ADDRESS'),
		STD_NAME_IN_WORDS: controller.get_attribute(student_info, 'STD_NAME_IN_WORDS'),
		STU_SESSION_RECORDS_TABLE: controller.get_attribute(student_info, 'STU_SESSION_RECORDS_TABLE'),
		HOUSE_NAME: controller.get_attribute(student_info, 'HOUSE_NAME'),
		CURRENT_CLASS_SECTION: controller.get_attribute(student_info, 'CURRENT_CLASS_SECTION'),
		HINDI_NAME: controller.get_attribute(student_info, 'HINDI_NAME'),
		MOTHER_HINDI: controller.get_attribute(student_info, 'MOTHER_HINDI'),
		FATHER_HINDI: controller.get_attribute(student_info, 'FATHER_HINDI'),
		NATIONALITY_HINDI: controller.get_attribute(student_info, 'NATIONALITY_HINDI'),
		DOJ_HINDI: controller.get_attribute(student_info, 'DOJ_HINDI'),
		JOINED_IN_CLASS_HINDI: controller.get_attribute(student_info, 'JOINED_IN_CLASS_HINDI'),
		DOB_HINDI: controller.get_attribute(student_info, 'DOB_HINDI'),
		DOB_WORDS_HINDI: controller.get_attribute(student_info, 'DOB_WORDS_HINDI'),
		CURRENT_CLASS_HINDI: controller.get_attribute(student_info, 'CURRENT_CLASS_HINDI'),
		LAST_EXAM_WITH_RESULT_HINDI: controller.get_attribute(student_info, 'LAST_EXAM_WITH_RESULT_HINDI'),
		FAILED_ONCE_OR_TWICE_IN_SAME_CLASS_HINDI: controller.get_attribute(student_info, 'FAILED_ONCE_OR_TWICE_IN_SAME_CLASS_HINDI'),
		SUBJECTS_STUDIED_HINDI: controller.get_attribute(student_info, 'SUBJECTS_STUDIED_HINDI'),
		PROMOTED_FOR_TC_HINDI: controller.get_attribute(student_info, 'PROMOTED_FOR_TC_HINDI'),
		FEES_PAID_UPTO_HINDI: controller.get_attribute(student_info, 'FEES_PAID_UPTO_HINDI'),
		FEE_CONCESSION_DETAILS_HINDI: controller.get_attribute(student_info, 'FEE_CONCESSION_DETAILS_HINDI'),
		TOTAL_WORKING_DAYS_HINDI: controller.get_attribute(student_info, 'TOTAL_WORKING_DAYS_HINDI'),
		TOTAL_PRESENT_DAYS_HINDI: controller.get_attribute(student_info, 'TOTAL_PRESENT_DAYS_HINDI'),
		NCC_CADET_DETAILS_HINDI: controller.get_attribute(student_info, 'NCC_CADET_DETAILS_HINDI'),
		GAMES_OR_EXTRACURRICULAR_ACT_HINDI: controller.get_attribute(student_info, 'GAMES_OR_EXTRACURRICULAR_ACT_HINDI'),
		ACHIEVEMENT_IN_GAMES_OR_EXTRACURRICULAR_ACT_HINDI: controller.get_attribute(student_info, 'ACHIEVEMENT_IN_GAMES_OR_EXTRACURRICULAR_ACT_HINDI'),
		CHARACTER_FOR_TC_HINDI: controller.get_attribute(student_info, 'CHARACTER_FOR_TC_HINDI'),
		DATE_OF_APPLICATION_FOR_TC_FORMATTED_HINDI: controller.get_attribute(student_info, 'DATE_OF_APPLICATION_FOR_TC_FORMATTED_HINDI'),
		TC_ISSUE_DATE_HINDI: controller.get_attribute(student_info, 'TC_ISSUE_DATE_HINDI'),
		REASON_FOR_LEAVING_HINDI: controller.get_attribute(student_info, 'REASON_FOR_LEAVING_HINDI'),
		ANY_OTHER_REMARKS_HINDI: controller.get_attribute(student_info, 'ANY_OTHER_REMARKS_HINDI'),
		CURRENT_ADDRESS_HINDI: controller.get_attribute(student_info, 'CURRENT_ADDRESS_HINDI'),
		DOL_HINDI: controller.get_attribute(student_info, 'DOL_HINDI'),
		ISSUED_TC_NUM_HINDI: controller.get_attribute(student_info, 'ISSUED_TC_NUM_HINDI'),
		BELONGS_TO_ST_SC_HINDI: controller.get_attribute(student_info, 'BELONGS_TO_ST_SC_HINDI'),
		MOTHER_TONGUE_NAME: controller.get_attribute(student_info, 'MOTHER_TONGUE_NAME'),
		STU_RELIGION: controller.get_attribute(student_info, 'RELIGION'),
		STU_CASTE: controller.get_attribute(student_info, 'CASTE'),
		STU_CASTE_CATEGORY_NAME: controller.get_attribute(student_info, 'CASTE_CATEGORY_NAME'),
		CURRENT_SESSION_YYYY: controller.get_attribute(student_info, 'CURRENT_SESSION_YYYY'),
		CURRENT_CLASS: controller.get_attribute(student_info, 'CURRENT_CLASS'),
		JOINED_IN_SESSION: controller.get_value(student_info.JOINED_IN_SESSION, "Joined in Session"),
		JOINED_IN_CLASS: controller.get_value(student_info.JOINED_IN_CLASS, "Joined in Class"),
		CERT_NUM: controller.get_value(student_info.CERT_NUM, "Cert Number"),
		DOB_WITH_MONTH_NAME: controller.get_value(student_info.DOB_WITH_MONTH_NAME, "Date of Birth"),
		PREVIOUS_SESSION: controller.get_value(student_info.PREVIOUS_SESSION, "Previous Session"),
		APAAR_ID: controller.get_value(student_info.APAAR_ID, "APAAR"),
		UDISE_PEN: controller.get_value(student_info.UDISE_NUMBER, "UDISE PEN"),
		MOTHER_OCCUPATION: controller.get_value(student_info.MOTHER_OCCUPATION, "Mother's Occupation"),
		PLACE_OF_BIRTH: controller.get_value(student_info.PLACE_OF_BIRTH, "Place of Birth"),
		POB_TALUKA: controller.get_value(student_info.POB_TALUKA, "Place of Birth Taluka"),
		POB_DIST: controller.get_value(student_info.POB_DIST, "Place of Birth Dist."),
		POB_STATE: controller.get_value(student_info.POB_STATE, "Place of Birth State"),
		FATHER_OCCUPATION: controller.get_value(student_info.FATHER_OCCUPATION, "Father's Occupation"),
		STU_UID_NUMBER: controller.get_value(student_info.STU_UID_NUMBER, "UID NUMBER"),
		STU_AADHAAR: controller.get_value(student_info.STU_AADHAAR, "Student's Aadhar Num"),
		BLOOD_GROUP: controller.get_value(student_info.BLOOD_GROUP, "Student's Blood Group"),
		MOTHER_AADHAAR: controller.get_value(student_info.MOTHER_AADHAAR, "Mother's Aadhar Num"),
		FATHER_AADHAAR: controller.get_value(student_info.FATHER_AADHAAR, "Father's Aadhar Num"),
		GUARDIAN: controller.get_value(student_info.GUARDIAN, "Guardian's Name"),

		MOTHER_OCCUPATION: controller.get_value(student_info.MOTHER_OCCUPATION, "Mother's Occupation"),
		MOTHER_QUAL_SCHOOLING: controller.get_value(student_info.MOTHER_QUAL_SCHOOLING, "Mother's Qualification"),
		FATHER_OCCUPATION: controller.get_value(student_info.FATHER_OCCUPATION, "Father's Occupation"),
		FATHER_QUAL_SCHOOLING: controller.get_value(student_info.FATHER_QUAL_SCHOOLING, "Father's Qualification"),

		FATHER_MOB: controller.get_value(student_info.FATHER_MOB, "Mother's Mobile"),
		MOTHER_MOB: controller.get_value(student_info.MOTHER_MOB, "Father's Mobile"),

		GENDER: controller.get_value(student_info.GENDER == 'M' ? 'Male' : 'Female', "Gender"),

		__TODAY__: today.toISOString().slice(0, 10).split('-').reverse().join('-'),
		__TODAY_YEAR__: today.getFullYear(),
		__TODAY_MONTH_NAME__: today.toLocaleString('default', { month: 'long' }),

		TODAY: today.toISOString().slice(0, 10).split('-').reverse().join('-'),
		TODAY_YEAR: today.getFullYear(),
		TODAY_MONTH_NAME: today.toLocaleString('default', { month: 'long' }),

		mst_or_miss: student_info.GENDER == 'M' ? 'Mst.' : 'Miss',
		son_or_daughter_of: student_info.GENDER == 'M' ? 'S/O' : 'D/O',
		he_or_she: student_info.GENDER == 'M' ? 'he' : 'she',
		son_or_daughter: student_info.GENDER == 'M' ? 'son' : 'daughter',
		He_or_She: student_info.GENDER == 'M' ? 'He' : 'She',
		his_or_her: student_info.GENDER == 'M' ? 'his' : 'her',
		him_or_her: student_info.GENDER == 'M' ? 'him' : 'her',
		His_or_Her: student_info.GENDER == 'M' ? 'His' : 'Her',

	});


	/*

	if (data4template.IMAGE_FILE_PATH == '') {
		$("span.STU_PIC").html('<span class="missing_value">IMAGE_FILE_PATH</span>');
	}
	else {
		$("span.STU_PIC").html('<img style="height: 150px" src="../../../../' + data4template.IMAGE_FILE_PATH + '">');
	}
	*/

	// Render the document (replace all placeholders)
    doc.render(data);

	replaceRedTags(doc);
	// Generate the updated document as a buffer
	const buf = doc.getZip().generate({
		type: 'nodebuffer',
		compression: 'DEFLATE',
	});
	const tempFile = data_directory + `temp_${temp_file_number}.docx`;
	await fs.writeFile(tempFile, buf);
	controller.tempWordFiles.push(tempFile);
	// Assuming getContentSize is an asynchronous function
	console.log(`Document Added! for ${student_info.STU_NAME}`);

	}
	catch (err) {
		console.log(err);
	}
};

// Getting page size
controller.getPageSize = async function (input_file_path) {
	// Read the DOCX file as a binary buffer
	const buffer = await fs.readFile(input_file_path);
	// console.log(buffer);

	// Load the DOCX file as a PizZip instance
	const zip = new PizZip(buffer);

	// Extract the document.xml file from the DOCX
	const documentXml = zip.file('word/document.xml').asText();

	// Parse XML content
	const result = await parseStringPromise(documentXml);

	// Extract page size information
	try {
		const sectPr = result['w:document']['w:body'][0]['w:sectPr'][0];
		const pgSz = sectPr['w:pgSz'][0]['$'];
		const width = parseInt(pgSz['w:w'], 10) / 20; // Convert twips to points
		const height = parseInt(pgSz['w:h'], 10) / 20; // Convert twips to points
		return { width, height };
	}
	catch (error) {
		throw new Error('Unable to extract page size: ' + error.message);
	}
};

// Getting size of contents
controller.getContentSize = async function(input_file_path) {
	// Read the DOCX file as a binary buffer
	const buffer = await fs.readFile(input_file_path);

	// Load the DOCX file as a PizZip instance
	const zip = new PizZip(buffer);

	// Extract the document.xml file from the DOCX
	const documentXml = zip.file('word/document.xml').asText();

	// Parse XML content
	const result = await parseStringPromise(documentXml);

	// Extract the text content
	let textContent = '';
	const paragraphs = result['w:document']['w:body'][0]['w:p'];

	for (const p of paragraphs) {
		if (p['w:r'] && p['w:r'][0]['w:t']) {
			textContent += p['w:r'][0]['w:t'].join(' ') + '\n';
		}
	}

	// Convert text content to plain text
	const plainText = htmlToText(textContent);

	// Estimate content size
	// Here we estimate size based on text length. This is a rough approximation.
	const textSizeEstimate = {
		charCount: plainText.length,
		// Approximate size in kilobytes
		sizeKb: Buffer.byteLength(plainText, 'utf8') / 1024
	};
	return textSizeEstimate;
};

// Estimating Dimension of content in word to fit contents
controller.estimateTextDimensions = function(charCount, fontSize = 22, charsPerLine = 10, linesPerPage = 100) {
	// Average width of a character in points
	const averageCharWidth = 0.6 * fontSize;
	// Line height including line spacing
	const lineHeight = fontSize * 1.2; // 20% extra for line spacing

	// Calculate number of lines needed
	const lines = Math.ceil(charCount / charsPerLine);

	// Calculate total height in points
	const totalHeightPoints = lines * lineHeight;

	// Convert points to inches (1 inch = 72 points)
	const totalHeightInches = totalHeightPoints / 72;

	// Calculate the width in inches assuming a fixed-width per line
	const totalWidthPoints = charsPerLine * averageCharWidth;
	const totalWidthInches = totalWidthPoints / 72;

	return {
		widthInches: totalWidthInches,
		heightInches: totalHeightInches,
		lines: lines
	};
};

// loading all temp files
controller.loadTempFile = async function (filePath) {
	return await fs.readFile(filePath);
};

// Merging Docx
controller.merging_doc = async function(data_directory, input_file_path, output_file_path, merge_small_certs_in_one_page) {
	let tempDocxBody = [];

	const OutputPageSize_in_points = await controller.getPageSize(input_file_path);
	const outputHeight_in_inches = OutputPageSize_in_points.height / 72;
	let outputPageHeight = outputHeight_in_inches;
	const contentSize_char = await controller.getContentSize(controller.tempWordFiles[0]);
	const dimensionsContentSize = controller.estimateTextDimensions(contentSize_char.charCount);

	const contentHeight = dimensionsContentSize.heightInches.toFixed(2);
	const pageBreak = `<w:p> <w:r> <w:br w:type="page"/> </w:r> </w:p>`;

	const files = [];
	for (let i = 0; i < controller.tempWordFiles.length; i++) {
		files.push(await controller.loadTempFile(controller.tempWordFiles[i]));
	}

	var docx = new DocxMerger({}, files.map((file) => {
		return file;
	}));
	// console.log(contentHeight, outputHeight_in_inches - 4, merge_small_certs_in_one_page);

	if (merge_small_certs_in_one_page) {
		if (contentHeight < outputHeight_in_inches - 4) {
			docx._body.forEach((content, index) => {
				if (index % 2 === 0) {
					tempDocxBody.push(content);
				}
			})
			docx._body = tempDocxBody;
			tempDocxBody = [];

			if (contentSize_char.charCount > 2) {
				docx._body.forEach((content, index) => {
					if (index % 2 === 0) {
						tempDocxBody.push(content);
					}
				});
				docx._body = tempDocxBody;
				tempDocxBody = [];
				docx._body.forEach((content, index) => {
					if (outputPageHeight - 1 < contentHeight) {
						tempDocxBody.push(pageBreak);
						tempDocxBody.push(content);
						outputPageHeight = outputHeight_in_inches;
					}
					else {
						outputPageHeight -= contentHeight;
						tempDocxBody.push(content);
					}
				});
				docx._body = tempDocxBody;
			}
		}
	}
	else {
		if (contentHeight > outputHeight_in_inches - 4) {
			docx._body.forEach((content, index) => {
				if (index % 2 === 0) {
					tempDocxBody.push(content);
				}
			})
			docx._body = tempDocxBody;
			tempDocxBody = [];
		}
	}

	let datas;
	// console.log(docx._body[0]);
	docx.save('nodebuffer', async function(data) {
		datas = data;
	})
	await fs.writeFile(output_file_path, datas, (err) => {
		if (err) {
			console.error("Error writing the output file:", err);
			rej();
		}
		else {
			// console.log("Document Created Successfully!!");
			res();
		}
	});
};

// Docx to Pdf
controller.convertDocxToPdf_officeTopdf = async function(output_file_path, output_file_path_pdf) {
	// console.log(output_file_path);
	var wordBuffer = await fs.readFile(output_file_path);
	await toPdf(wordBuffer).then(
		async (pdfBuffer) => {
			await fs.writeFile(output_file_path_pdf, pdfBuffer);
		},
		(err) => {
			console.log(err);
		}
	);
};


controller.convertDocxToPdf_libre = async function (output_file_path, output_file_path_pdf) {
	const ext = '.pdf'
	const docxBuf = await fs.readFile(output_file_path);
	// console.log(docxBuf)

	const pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
	// Here in done you have pdf file which you can save or transfer in another stream
	await fs.writeFile(output_file_path_pdf, pdfBuf);
}

module.exports = controller.run;

import {PageGeneratorRedirectBase} from "./generate-pages-page-generator.js";

class _PageGeneratorFaleriaTools extends PageGeneratorRedirectBase {
	_page = "faleriatools.html";

	_pageDescription = "A suite of tools for the Adventurers of Faleria to use.";

	_redirectHref = "index.html";
	_redirectMessage = "the homepage";
}

export const PAGE_GENERATORS_REDIRECT = [
	new _PageGeneratorFaleriaTools(),
];

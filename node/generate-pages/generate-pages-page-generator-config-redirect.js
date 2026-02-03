import {PageGeneratorRedirectBase} from "./generate-pages-page-generator.js";

class _PageGenerator5etools extends PageGeneratorRedirectBase {
	_page = "5etools.html";

	_pageDescription = "A suite of tools for the Adventurers of Faleria to use.";

	_redirectHref = "index.html";
	_redirectMessage = "the homepage";
}

export const PAGE_GENERATORS_REDIRECT = [
	new _PageGenerator5etools(),
];

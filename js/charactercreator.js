import {StatGenUi} from "./statgen/statgen-ui.js";

class CharacterCreatorPage {
	constructor () {
		this._state = {
			ixRace: null,
			ixBackground: null,
			ixClass: null,
			ixSubclass: null,
			ixFeat: null,
			activeStep: "overview",
			searchRace: "",
			searchBackground: "",
			searchClass: "",
			searchFeat: "",
		};

		this._races = [];
		this._backgrounds = [];
		this._classes = [];
		this._feats = [];
		this._statGenUi = null;
		this._$statGenHost = null;
	}

	async pInit () {
		await Promise.all([
			PrereleaseUtil.pInit(),
			BrewUtil2.pInit(),
		]);
		await ExcludeUtil.pInitialise();

		const [races, backgrounds, classes, feats] = await Promise.all([
			this._pLoadRaces(),
			this._pLoadBackgrounds(),
			this._pLoadClasses(),
			this._pLoadFeats(),
		]);
		this._races = races.sort((a, b) => SortUtil.ascSortLower(a.name, b.name));
		this._backgrounds = backgrounds.sort((a, b) => SortUtil.ascSortLower(a.name, b.name));
		this._classes = classes.sort((a, b) => SortUtil.ascSortLower(a.name, b.name));
		this._feats = feats.sort((a, b) => SortUtil.ascSortLower(a.name, b.name));

		this._statGenUi = new StatGenUi({
			races: this._races,
			backgrounds: this._backgrounds,
			feats: this._feats,
			isCharacterMode: true,
		});
		await this._statGenUi.pInit();
		const ixPointBuy = this._statGenUi.MODES.indexOf("pointbuy");
		if (~ixPointBuy) this._statGenUi.ixActiveTab = ixPointBuy;
		this._$statGenHost = es(`<div></div>`);
		this._statGenUi.render(this._$statGenHost);

		this._render();
		window.dispatchEvent(new Event("toolsLoaded"));
	}

	async _pLoadRaces () {
		return [
			...(await DataUtil.race.loadJSON()).race,
			...((await DataUtil.race.loadPrerelease({isAddBaseRaces: false})).race || []),
			...((await DataUtil.race.loadBrew({isAddBaseRaces: false})).race || []),
		]
			.filter(it => {
				const hash = UrlUtil.URL_TO_HASH_BUILDER[UrlUtil.PG_RACES](it);
				return !ExcludeUtil.isExcluded(hash, "race", it.source);
			});
	}

	async _pLoadBackgrounds () {
		return [
			...(await DataUtil.loadJSON("data/backgrounds.json")).background,
			...((await PrereleaseUtil.pGetBrewProcessed()).background || []),
			...((await BrewUtil2.pGetBrewProcessed()).background || []),
		]
			.filter(it => {
				const hash = UrlUtil.URL_TO_HASH_BUILDER[UrlUtil.PG_BACKGROUNDS](it);
				return !ExcludeUtil.isExcluded(hash, "background", it.source);
			});
	}

	async _pLoadClasses () {
		return [
			...(await DataUtil.class.loadJSON()).class,
			...((await DataUtil.class.loadPrerelease()).class || []),
			...((await DataUtil.class.loadBrew()).class || []),
		]
			.filter(it => {
				const hash = UrlUtil.URL_TO_HASH_BUILDER[UrlUtil.PG_CLASSES](it);
				return !ExcludeUtil.isExcluded(hash, "class", it.source);
			});
	}

	async _pLoadFeats () {
		return [
			...(await DataUtil.loadJSON("data/feats.json")).feat,
			...((await PrereleaseUtil.pGetBrewProcessed()).feat || []),
			...((await BrewUtil2.pGetBrewProcessed()).feat || []),
		]
			.filter(it => {
				const hash = UrlUtil.URL_TO_HASH_BUILDER[UrlUtil.PG_FEATS](it);
				return !ExcludeUtil.isExcluded(hash, "feat", it.source);
			});
	}

	_render () {
		const $root = es("#character-creator-main").empty();

		const steps = [
			{id: "overview", title: "Overview", sub: "Starter guidance"},
			{id: "race", title: "Race", sub: "Select ancestry"},
			{id: "class", title: "Class", sub: "Class and subclass"},
			{id: "background", title: "Background", sub: "Origin and skills"},
			{id: "feat", title: "Feat", sub: "Optional feat choice"},
			{id: "stats", title: "Ability Scores", sub: "Point Buy"},
			{id: "summary", title: "Summary", sub: "Review build"},
		];

		const $steps = es(`<div class="character-creator__steps"></div>`);
		steps.forEach(step => {
			const isActive = this._state.activeStep === step.id;
			const $btn = es(`<button class="ve-btn ${isActive ? "ve-btn-primary" : "ve-btn-default"} character-creator__step"></button>`)
				.on("click", () => {
					this._state.activeStep = step.id;
					this._render();
				});
			$btn.append(es(`<div class="character-creator__step-title">${step.title}</div>`));
			$btn.append(es(`<div class="character-creator__step-sub">${step.sub}</div>`));
			$steps.append($btn);
		});

		const $panel = es(`<div class="character-creator__panel"></div>`);
		this._renderActivePanel($panel);

		const $layout = es(`<div class="character-creator__layout"></div>`);
		$layout.append($steps);
		$layout.append($panel);
		$root.append($layout);
	}

	_renderActivePanel ($panel) {
		switch (this._state.activeStep) {
			case "race": return this._renderRaceStep($panel);
			case "class": return this._renderClassStep($panel);
			case "background": return this._renderBackgroundStep($panel);
			case "feat": return this._renderFeatStep($panel);
			case "stats": return this._renderStatsStep($panel);
			case "summary": return this._renderSummaryStep($panel);
			case "overview":
			default: return this._renderOverviewStep($panel);
		}
	}

	_renderOverviewStep ($panel) {
		$panel.append(es(`<h3 class="mt-0">Character Creator</h3>`));
		$panel.append(es(`<p>Build your character in guided steps inspired by modern VTT creators. This flow uses site content data for races, classes, backgrounds, and feats.</p>`));
		$panel.append(es(`<p><b>Stat Allocation</b> is powered by the built-in Point Buy system on the Ability Scores step.</p>`));
		$panel.append(es(`<p>Use the step list on the left to progress and revisit choices at any time.</p>`));
	}

	_renderRaceStep ($panel) {
		const selected = this._state.ixRace == null ? null : this._races[this._state.ixRace];
		$panel.append(es(`<h3 class="mt-0">Choose a Race</h3>`));
		$panel.append(this._$getSearchBox({
			prop: "searchRace",
			placeholder: "Search races...",
		}));

		const search = (this._state.searchRace || "").trim().toLowerCase();
		const races = this._races.filter(it => !search || it.name.toLowerCase().includes(search));
		const $list = es(`<div class="character-creator__list"></div>`);
		races.forEach((race, ix) => {
			const actualIx = this._races.indexOf(race);
			const isSel = actualIx === this._state.ixRace;
			const $card = es(`<div class="character-creator__card"></div>`);
			$card.append(es(`<div><b>${race.name}</b></div>`));
			$card.append(es(`<div class="ve-muted ve-small">${Parser.sourceJsonToAbv(race.source)}</div>`));
			$card.append(es(`<button class="ve-btn ve-btn-xs ${isSel ? "ve-btn-primary" : "ve-btn-default"} mt-2">${isSel ? "Selected" : "Select"}</button>`)
				.on("click", () => {
					this._state.ixRace = actualIx;
					this._statGenUi.ixRace = actualIx;
					this._render();
				}));
			$list.append($card);
		});
		$panel.append($list);

		if (selected) {
			$panel.append(es(`<div class="character-creator__summary"><b>Selected:</b> ${selected.name}</div>`));
		}
	}

	_renderClassStep ($panel) {
		const cls = this._state.ixClass == null ? null : this._classes[this._state.ixClass];
		$panel.append(es(`<h3 class="mt-0">Choose a Class</h3>`));
		$panel.append(this._$getSearchBox({
			prop: "searchClass",
			placeholder: "Search classes...",
		}));

		const search = (this._state.searchClass || "").trim().toLowerCase();
		const classes = this._classes.filter(it => !search || it.name.toLowerCase().includes(search));
		const $list = es(`<div class="character-creator__list"></div>`);
		classes.forEach(c => {
			const actualIx = this._classes.indexOf(c);
			const isSel = actualIx === this._state.ixClass;
			const $card = es(`<div class="character-creator__card"></div>`);
			$card.append(es(`<div><b>${c.name}</b></div>`));
			$card.append(es(`<div class="ve-muted ve-small">${Parser.sourceJsonToAbv(c.source)}</div>`));
			$card.append(es(`<button class="ve-btn ve-btn-xs ${isSel ? "ve-btn-primary" : "ve-btn-default"} mt-2">${isSel ? "Selected" : "Select"}</button>`)
				.on("click", () => {
					this._state.ixClass = actualIx;
					this._state.ixSubclass = null;
					this._render();
				}));
			$list.append($card);
		});
		$panel.append($list);

		if (cls?.subclasses?.length) {
			$panel.append(es(`<hr class="hr-3">`));
			$panel.append(es(`<h4>Subclass</h4>`));
			const subclasses = cls.subclasses;
			const $subWrp = es(`<div class="character-creator__list"></div>`);
			subclasses.forEach(sc => {
				const ixSc = cls.subclasses.indexOf(sc);
				const isSel = ixSc === this._state.ixSubclass;
				const $card = es(`<div class="character-creator__card"></div>`);
				$card.append(es(`<div><b>${sc.name}</b></div>`));
				$card.append(es(`<div class="ve-muted ve-small">${Parser.sourceJsonToAbv(sc.source)}</div>`));
				$card.append(es(`<button class="ve-btn ve-btn-xs ${isSel ? "ve-btn-primary" : "ve-btn-default"} mt-2">${isSel ? "Selected" : "Select"}</button>`)
					.on("click", () => {
						this._state.ixSubclass = ixSc;
						this._render();
					}));
				$subWrp.append($card);
			});
			$panel.append($subWrp);
		}
	}

	_renderBackgroundStep ($panel) {
		const selected = this._state.ixBackground == null ? null : this._backgrounds[this._state.ixBackground];
		$panel.append(es(`<h3 class="mt-0">Choose a Background</h3>`));
		$panel.append(this._$getSearchBox({
			prop: "searchBackground",
			placeholder: "Search backgrounds...",
		}));

		const search = (this._state.searchBackground || "").trim().toLowerCase();
		const backgrounds = this._backgrounds.filter(it => !search || it.name.toLowerCase().includes(search));
		const $list = es(`<div class="character-creator__list"></div>`);
		backgrounds.forEach(bg => {
			const actualIx = this._backgrounds.indexOf(bg);
			const isSel = actualIx === this._state.ixBackground;
			const $card = es(`<div class="character-creator__card"></div>`);
			$card.append(es(`<div><b>${bg.name}</b></div>`));
			$card.append(es(`<div class="ve-muted ve-small">${Parser.sourceJsonToAbv(bg.source)}</div>`));
			$card.append(es(`<button class="ve-btn ve-btn-xs ${isSel ? "ve-btn-primary" : "ve-btn-default"} mt-2">${isSel ? "Selected" : "Select"}</button>`)
				.on("click", () => {
					this._state.ixBackground = actualIx;
					this._statGenUi.ixBackground = actualIx;
					this._render();
				}));
			$list.append($card);
		});
		$panel.append($list);

		if (selected) $panel.append(es(`<div class="character-creator__summary"><b>Selected:</b> ${selected.name}</div>`));
	}

	_renderFeatStep ($panel) {
		const selected = this._state.ixFeat == null ? null : this._feats[this._state.ixFeat];
		$panel.append(es(`<h3 class="mt-0">Choose a Feat (Optional)</h3>`));
		const $top = es(`<div class="ve-flex-v-center mb-2"></div>`);
		$top.append(this._$getSearchBox({
			prop: "searchFeat",
			placeholder: "Search feats...",
		}));
		$top.append(es(`<button class="ve-btn ve-btn-xs ve-btn-default ml-2">Clear</button>`)
			.on("click", () => {
				this._state.ixFeat = null;
				this._render();
			}));
		$panel.append($top);

		const search = (this._state.searchFeat || "").trim().toLowerCase();
		const feats = this._feats.filter(it => !search || it.name.toLowerCase().includes(search));
		const $list = es(`<div class="character-creator__list"></div>`);
		feats.forEach(ft => {
			const actualIx = this._feats.indexOf(ft);
			const isSel = actualIx === this._state.ixFeat;
			const $card = es(`<div class="character-creator__card"></div>`);
			$card.append(es(`<div><b>${ft.name}</b></div>`));
			$card.append(es(`<div class="ve-muted ve-small">${Parser.sourceJsonToAbv(ft.source)}</div>`));
			$card.append(es(`<button class="ve-btn ve-btn-xs ${isSel ? "ve-btn-primary" : "ve-btn-default"} mt-2">${isSel ? "Selected" : "Select"}</button>`)
				.on("click", () => {
					this._state.ixFeat = actualIx;
					this._render();
				}));
			$list.append($card);
		});
		$panel.append($list);

		if (selected) $panel.append(es(`<div class="character-creator__summary"><b>Selected:</b> ${selected.name}</div>`));
	}

	_renderStatsStep ($panel) {
		$panel.append(es(`<h3 class="mt-0">Ability Scores (Point Buy)</h3>`));
		$panel.append(es(`<p class="ve-muted">This section uses the built-in Stat Generator Point Buy system.</p>`));

		const ixPointBuy = this._statGenUi.MODES.indexOf("pointbuy");
		if (~ixPointBuy) this._statGenUi.ixActiveTab = ixPointBuy;

		$panel.append(this._$statGenHost);
	}

	_renderSummaryStep ($panel) {
		const race = this._state.ixRace == null ? null : this._races[this._state.ixRace];
		const bg = this._state.ixBackground == null ? null : this._backgrounds[this._state.ixBackground];
		const cls = this._state.ixClass == null ? null : this._classes[this._state.ixClass];
		const sc = cls && this._state.ixSubclass != null ? cls.subclasses[this._state.ixSubclass] : null;
		const feat = this._state.ixFeat == null ? null : this._feats[this._state.ixFeat];
		const totals = this._statGenUi.getTotals();
		const pbTotals = totals?.totals?.pointbuy || {};

		$panel.append(es(`<h3 class="mt-0">Character Summary</h3>`));
		$panel.append(es(`<div class="mb-1"><b>Race:</b> ${race ? `${race.name} (${Parser.sourceJsonToAbv(race.source)})` : "Not selected"}</div>`));
		$panel.append(es(`<div class="mb-1"><b>Class:</b> ${cls ? `${cls.name} (${Parser.sourceJsonToAbv(cls.source)})` : "Not selected"}</div>`));
		$panel.append(es(`<div class="mb-1"><b>Subclass:</b> ${sc ? `${sc.name} (${Parser.sourceJsonToAbv(sc.source)})` : "Not selected"}</div>`));
		$panel.append(es(`<div class="mb-1"><b>Background:</b> ${bg ? `${bg.name} (${Parser.sourceJsonToAbv(bg.source)})` : "Not selected"}</div>`));
		$panel.append(es(`<div class="mb-3"><b>Feat:</b> ${feat ? `${feat.name} (${Parser.sourceJsonToAbv(feat.source)})` : "Not selected"}</div>`));

		$panel.append(es(`<h4>Ability Scores (Point Buy)</h4>`));
		$panel.append(es(`<div><b>STR</b> ${pbTotals.str ?? "\u2014"} | <b>DEX</b> ${pbTotals.dex ?? "\u2014"} | <b>CON</b> ${pbTotals.con ?? "\u2014"} | <b>INT</b> ${pbTotals.int ?? "\u2014"} | <b>WIS</b> ${pbTotals.wis ?? "\u2014"} | <b>CHA</b> ${pbTotals.cha ?? "\u2014"}</div>`));
	}

	_$getSearchBox ({prop, placeholder}) {
		return es(`<input class="form-control mb-2" type="search" placeholder="${placeholder}">`)
			.val(this._state[prop] || "")
			.on("input", evt => {
				this._state[prop] = evt.currentTarget.value;
				this._render();
			});
	}
}

const characterCreatorPage = new CharacterCreatorPage();
window.addEventListener("load", () => characterCreatorPage.pInit());

import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting } from 'obsidian';
import { Perfume } from 'perfume';
const XLSX = require('xlsx');
const fs = require('fs');

interface MyPluginSettings {
	mySetting: string;
	excelFilePath: string;
	keywordSheetName: string;
	accordSheetName: string;
	outputDirectory: string;
}

const DEFAULT_SETTINGS: MyPluginSettings = {
	mySetting: 'default',
	excelFilePath: '',
	keywordSheetName: '',
	accordSheetName: '',
	outputDirectory: '',
}

export default class MyPlugin extends Plugin {
	settings: MyPluginSettings;
	perfumeMap: Map<string, Perfume>;
	accordSet: Set<string>;

	onSubmit = async (): Promise<void> => {
		const statusBarItemEl = this.addStatusBarItem();
		statusBarItemEl.setText(`[Etov Working🚀]`);
		// Todo.필요한 시트를 열되, 안열렸을때 파일이없을때 예외처리가 필요하다.
		// open excelfile
		const filepath = this.settings.excelFilePath;
		const excelFile = fs.readFileSync(filepath);
		const workbook = XLSX.read(excelFile, { type: 'buffer' });

		// Get the first sheet as JSON
		const keyWordSheet = workbook.Sheets[this.settings.keywordSheetName];
		const keyWordJsonData = XLSX.utils.sheet_to_json(keyWordSheet, {header:1});

		const accordSheet = workbook.Sheets[this.settings.accordSheetName];
		const accordJsonData: (string)[][] = XLSX.utils.sheet_to_json(accordSheet, {header:1});
		// 비어있는 셀도 같이 잡혀서 필터링함.
		const filteredAccordData = accordJsonData.filter(row => row.some(cell => cell !== null && cell !== undefined && cell !== ''));

		// 파일 read 시작
		for (let index = 1; index < keyWordJsonData.length; index++) {
			const key = keyWordJsonData[index][0];
			if (this.perfumeMap.has(key)) {
				this.perfumeMap.get(key)?.keywords.push(keyWordJsonData[index][3]);
			} else {
				this.perfumeMap.set(key, new Perfume(keyWordJsonData[index][0], keyWordJsonData[index][1], keyWordJsonData[index][2], [keyWordJsonData[index][3]]));
			}
		}

		for (let index = 1; index < filteredAccordData.length; index++) {
			const key = filteredAccordData[index][0];
			if (this.perfumeMap.has(key)) {
				if (filteredAccordData[index][3] == undefined) continue;
				const accordsArray = (filteredAccordData[index][3] as string).split(',');
				const trimmedAccords = accordsArray.map(accord => accord.trim());

				const perfume = this.perfumeMap.get(key);
				if (perfume && perfume.accords) perfume.accords.push(...trimmedAccords);
				// trimmedAccords의 각 원소를
			} else {
				console.log(`❌ key: ${key} | name: ${filteredAccordData[index][2]}에 해당하는 향수가 없습니다...!`)
			}
		}
		// 파일 read 끝
		console.log('perfumeMap:', this.perfumeMap);
		await this.createPerfumeFiles();
		statusBarItemEl.setText(`[Etov Done ✅]`);
	}

	async onload() {
		await this.loadSettings();
		this.perfumeMap = new Map<string, Perfume>();
		this.accordSet = new Set<string>();

		const ribbonIconEl = this.addRibbonIcon('dice', 'Sample Plugin', (evt: MouseEvent) => {
			new Notice('This is a notice!');
			this.onSubmit();
			//new SampleModal(this.app, this.onSubmit).open();
			// 엑셀파일 입력 모달을 open한다.
		});
		ribbonIconEl.addClass('my-plugin-ribbon-class');

		// This adds a status bar item to the bottom of the app. Does not work on mobile apps.
		const statusBarItemEl = this.addStatusBarItem();
		statusBarItemEl.setText(`[Etov Online🌈]`);

		// This adds a simple command that can be triggered anywhere
		this.addCommand({
			id: 'open-sample-modal-simple',
			name: 'Open sample modal (simple)',
			callback: () => {
				new SampleModal(this.app, this.onSubmit).open();
			}
		});
		// This adds an editor command that can perform some operation on the current editor instance
		this.addCommand({
			id: 'sample-editor-command',
			name: 'Sample editor command',
			editorCallback: (editor: Editor, view: MarkdownView) => {
				console.log(editor.getSelection());
				editor.replaceSelection('Sample Editor Command');
			}
		});
		// This adds a complex command that can check whether the current state of the app allows execution of the command
		this.addCommand({
			id: 'open-sample-modal-complex',
			name: 'Open sample modal (complex)',
			checkCallback: (checking: boolean) => {
				// Conditions to check
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView) {
					// If checking is true, we're simply "checking" if the command can be run.
					// If checking is false, then we want to actually perform the operation.
					if (!checking) {
						new SampleModal(this.app, this.onSubmit).open();
					}

					// This command will only show up in Command Palette when the check function returns true
					return true;
				}
			}
		});

		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new SampleSettingTab(this.app, this));

		// If the plugin hooks up any global DOM events (on parts of the app that doesn't belong to this plugin)
		// Using this function will automatically remove the event listener when this plugin is disabled.
		// this.registerDomEvent(document, 'click', (evt: MouseEvent) => {
		// 	console.log('click', evt);
		// });

		// When registering intervals, this function will automatically clear the interval when the plugin is disabled.
		//this.registerInterval(window.setInterval(() => console.log('setInterval'), 5 * 60 * 1000));
	}

	onunload() {

	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	async createPerfumeFiles() {
		// 파일 생성 시작
		const outputDirectory = this.settings.outputDirectory
		let vault = this.app.vault;
		let keys = Array.from(this.perfumeMap.keys());

		// Ensure output directory exists
		let folder = vault.getAbstractFileByPath(outputDirectory);
		console.log('outputDirectory:', outputDirectory);
		console.log('folder:', folder);
		if (!folder) {
			await vault.createFolder(outputDirectory);
			await vault.createFolder(outputDirectory + '/perfume');
			await vault.createFolder(outputDirectory + '/accord');
		}

		for(let key of keys) {
			let perfume = this.perfumeMap.get(key);
			let fileName = `${outputDirectory}/perfume/${perfume?.name}.md`;
			// Format keywords with hashtags
			const hashtagKeywords = perfume?.keywords.map(keyword => `#${keyword}`).join('\n');
			let content = `# 향수명: ${perfume?.name}\n\n- 브랜드: [[${perfume?.brandName}]]\n- 키워드: ${hashtagKeywords}\n- 어코드:`;

			// perfume.accords의 각 원소들에 대해 	accordSet: Set<string>에 추가한다.
			if (perfume?.accords) {
				for (let accord of perfume.accords) {
					this.accordSet.add(accord);
					content += `\n[[${accord}]]`;
				}
			}

			if (!vault.getAbstractFileByPath(fileName))
				await vault.create(fileName, content);
			else
				await vault.adapter.write(fileName, content);
		}

		for(const accord of this.accordSet) {
			await vault.create(`${outputDirectory}/accord/${accord}.md`, '');
		}

		// 파일 생성 끝
	}
}

class SampleModal extends Modal {
	result: string = "/Users/simgeon-u/Desktop/Study/키워드-향수.xlsx";
  onSubmit: (result: string) => void;

	constructor(app: App, onSubmit: (result: string) => void) {
		super(app);
		this.onSubmit = onSubmit;
	}

	onOpen() {
		const { contentEl } = this;

    contentEl.createEl("h1", { text: "Excel 파일을 입력하세요 🔍" });

    new Setting(contentEl)
      .setName("Excel file")
      .addText((text) =>
        text.onChange((value) => {
          this.result = value
        }));

    new Setting(contentEl)
      .addButton((btn) =>
        btn
          .setButtonText("실행")
          .setCta()
          .onClick(() => {
            this.close();
            this.onSubmit(this.result);
          }));
	}

	onClose() {
		const {contentEl} = this;
		contentEl.empty();
	}
}

class SampleSettingTab extends PluginSettingTab {
	plugin: MyPlugin;

	constructor(app: App, plugin: MyPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const {containerEl} = this;

		containerEl.empty();

		containerEl.createEl('h2', {text: 'Settings for my awesome plugin.'});

		new Setting(containerEl)
      .setName('Excel File Path')
      .setDesc('Specify the path of the Excel file')
      .addText(text => text
        .setPlaceholder('Enter the path of the Excel file')
        .setValue(this.plugin.settings.excelFilePath || '')
        .onChange(async (value) => {
          this.plugin.settings.excelFilePath = value;
          await this.plugin.saveSettings();
        }));

		new Setting(containerEl)
      .setName('Keyword Sheet Name')
      .setDesc('Specify the name of the Keyword Sheet')
      .addText(text => text
        .setPlaceholder('Enter the name of the Keyword Sheet')
        .setValue(this.plugin.settings.keywordSheetName || '')
        .onChange(async (value) => {
          this.plugin.settings.keywordSheetName = value;
          await this.plugin.saveSettings();
        }));

		new Setting(containerEl)
      .setName('Accord Sheet Name')
      .setDesc('Specify the name of the Accord Sheet')
      .addText(text => text
        .setPlaceholder('Enter the name of the Accord Sheet')
        .setValue(this.plugin.settings.accordSheetName || '')
        .onChange(async (value) => {
          this.plugin.settings.accordSheetName = value;
          await this.plugin.saveSettings();
        }));

		new Setting(containerEl)
    .setName('Output Directory')
    .setDesc('Directory for the output files')
    .addText(text => text
        .setPlaceholder('Enter your directory')
        .setValue(this.plugin.settings.outputDirectory)
        .onChange(async(value) => {
            this.plugin.settings.outputDirectory = value;
            await this.plugin.saveSettings();
        }));
	}
}

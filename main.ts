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

	onSubmit = async (): Promise<void> => {
		// open excelfile
		const filepath = this.settings.excelFilePath;
		const excelFile = fs.readFileSync(filepath);
		const workbook = XLSX.read(excelFile, { type: 'buffer' });

		// Get the first sheet as JSON
		//const sheetName = workbook.SheetNames[0];
		const sheetName = '2차 정리';
		const worksheet = workbook.Sheets[sheetName];
		const jsonData = XLSX.utils.sheet_to_json(worksheet, {header:1});


		//let perfumeMap = new Map<string, Perfume>();
		for (let index = 1; index < jsonData.length; index++) {
		// for (let index = 1; index < 100; index++) {
			const key = jsonData[index][0];

			if (this.perfumeMap.has(key)) {
				this.perfumeMap.get(key)?.keywords.push(jsonData[index][3]);
			} else {
				this.perfumeMap.set(key, new Perfume(jsonData[index][0], jsonData[index][1], jsonData[index][2], [jsonData[index][3]]));
			}
		}
		console.log('perfumeMap:', this.perfumeMap);
		await this.createPerfumeFiles();
	}

	async onload() {
		await this.loadSettings();
		this.perfumeMap = new Map<string, Perfume>();

		const ribbonIconEl = this.addRibbonIcon('dice', 'Sample Plugin', (evt: MouseEvent) => {
			new Notice('This is a notice!');
			this.onSubmit();
			//new SampleModal(this.app, this.onSubmit).open();
			// 엑셀파일 입력 모달을 open한다.
		});
		ribbonIconEl.addClass('my-plugin-ribbon-class');

		// This adds a status bar item to the bottom of the app. Does not work on mobile apps.
		const statusBarItemEl = this.addStatusBarItem();
		statusBarItemEl.setText(`Etov online! ⚙debug⚙: |${this.settings.keywordSheetName}|${this.settings.accordSheetName}||`);

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
		const outputDirectory = this.settings.outputDirectory
		let vault = this.app.vault;
		let keys = Array.from(this.perfumeMap.keys());

		// Ensure output directory exists
		let folder = vault.getAbstractFileByPath(outputDirectory);
		console.log('outputDirectory:', outputDirectory);
		console.log('folder:', folder);
		if (!folder) {
			await vault.createFolder(outputDirectory);
		}

		for(let key of keys) {
			let perfume = this.perfumeMap.get(key);
			let fileName = `${outputDirectory}/${perfume?.name}.md`;
			// Format keywords with hashtags
			const hashtagKeywords = perfume?.keywords.map(keyword => `#${keyword}`).join('\n');
			const content = `# 향수명: ${perfume?.name}\n\n- 브랜드: [[${perfume?.brandName}]]\n- 키워드: ${hashtagKeywords}`;

			if (!vault.getAbstractFileByPath(fileName))
				//await vault.create(fileName, '', true);
				await vault.create(fileName, content);
			else
				await vault.adapter.write(fileName, content);
		}
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

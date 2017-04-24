var styles = {
	header : {
		font : {
			bold: true,
			size: 16
		},
		fill : {
			type: "pattern",
			pattern:"solid",
			fgColor:{argb:"D9D7DC"}
		},
		border : {
			top: {style:"thin", color: {argb:"353535"}},
			left: {style:"thin", color: {argb:"353535"}},
			bottom: {style:"thin", color: {argb:"353535"}},
			right: {style:"thin", color: {argb:"353535"}}
		}
	},
	cellKey : {
		alignment : {
			wrapText : true,
			vertical: "top"
		},
		font : {
			color: { argb: "2600FF" },
			size: 12,
			name: "Arial",
			family: 2
		}
	},
	cellTranslation : {
		alignment : {
			wrapText : true,
			vertical: "top"
		},
		font : {
			size: 12,
			name: "Arial",
			family: 2
		}
	},
	emptyRow: {
		fill : {
			type: "pattern",
			pattern:"solid",
			fgColor:{argb:"7DBB52"}
		}
	},
};

var config = {
	baseColumns: [
		{
			header: "PAGE",
			key: "page",
			width: 22 ,
			style: styles.cellTranslation
		},
		{
			header: "SECTION",
			key: "section",
			width: 30 ,
			style: styles.cellTranslation
		},
		{
			header: "KEY",
			key: "key",
			width: 40 ,
			style: styles.cellKey
		}
	]
};

var optionalConfigFileName = 'translationConfig.json';

module.exports = {
	config: config,
	styles: styles,
	optionalConfigFileName: optionalConfigFileName
}

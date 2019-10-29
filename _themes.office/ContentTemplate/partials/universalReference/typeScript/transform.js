var isNode = typeof process !== 'undefined';
var contentCommon = require(isNode ? '../../../content.common.js' : './content.common.js');

function formatTypeStrings(model, type) {
	var objNotationRegEx = /: \@([^),]+)/g,
		typeParamRegEx = /^([^<]+)<([^>]+)>$/,
		uidStr = null,
		uidMatches = [];

	//check for inlined types
	if (type) {
		for (var k = 0; k < type.length; k++) {
			uidMatches = [];
			uidStr = type[k].uid;
			type[k].postType = '';

			if (uidStr.indexOf('{') === 0) {
				// object notation
				// example: { new (serviceScope: @sp-core-library.ServiceScope); }
				match = objNotationRegEx.exec(uidStr);

				while (match != null) {
					uidMatches.push(match[1]);
					match = objNotationRegEx.exec(uidStr);
				}

				if (uidMatches.length) {
					for (var k2 = 0; k2 < uidMatches.length; k2++) {
						uidStr = uidStr.replace(uidMatches[k2], '<xref uid=\"' + uidMatches[k2] + '\" />');
					}

					uidStr = uidStr.replace('@', '');

					if (type[k].specName.length) {
						type[k].specName[0].value = uidStr;
					}
				}
			} else if (type[k].uid.match(typeParamRegEx)) {
				// typed param
				// example: @sp-webpart-base.IPropertyPaneField<@sp-webpart-base.BaseClientSideWebPart>
				match = typeParamRegEx.exec(uidStr);
				if (match.length && match[1] && match[2] && type[k].specName.length) {

					var postType = '';
					if (match[2].indexOf('[]') > -1) {
						match[2] = match[2].replace('[]', '');
						postType = '[]';
					}

					type[k].specName[0].value = '<xref uid=\"' + match[1] + '\" />&lt;<xref uid=\"' + match[2] + '\" />' + postType + '&gt;';
				}
			} else {
				//basic uid
				// remove '[]' if present
				if (type[k].specName && type[k].specName.length) {
					if (type[k].specName[0].value.indexOf('[]') > -1) {
						type[k].specName[0].value = type[k].specName[0].value.replace('[]', '');
						type[k].postType = '[]';
					}
				}
			}
		}
	}
}

function updateMembers(model, members) {
	var newMembers = [];

	if (members) {
		var m = null;

		for (var i = 0; i < members.length; i++) {
			m = members[i];
			if (m.name && m.name[0] && m.name[0].value) {
				m.id = contentCommon.createHtmlId(m.name[0].value);
			} else {
				m.id = contentCommon.createHtmlId(m.uid);
			}

			if (m.deprecated) {
				if (m.deprecated.content) {
					m.deprecated.content = contentCommon.parseMarkdown(m.deprecated.content, model._key);
				}
			} else {
				m.deprecated = null;
			}

			if (!m.isPreview) {
				m.isPreview = false;
			}

			if (!m.summary) {
				m.summary = null;
			} else {
				m.summary = contentCommon.parseMarkdown(m.summary, model._key);
			}

			if (m.remarks === undefined) {
				m.remarks = null;
			}

			if (m.name && m.name[0] && m.name[0].value) {
				m.name[0].value = contentCommon.breakText(m.name[0].value.replace('<', '&lt;').replace('>', '&gt;'));
			}

			updateParameters(model, m);

			if (m.syntax && m.syntax.return) {
				for (var j = 0; j < m.syntax.return.length; j++) {
					if (m.syntax.return[j].value && m.syntax.return[j].value.type) {
						formatTypeStrings(model, m.syntax.return[j].value.type);
						for (var k = 0; k < m.syntax.return[j].value.type.length; k++) {
							if (k > 0) {
								m.syntax.return[j].value.type[k].className = " halfStack";
							}
						}
					}
				}
			}

			if (m.exceptions && m.exceptions[0] && m.exceptions[0].value) {
				for (var j = 0; j < m.exceptions[0].value.length; j++) {
					if (j > 0) {
						m.exceptions[0].value[j].className = " halfStack";
					}
				}
			}

			if (m.type === "property") {
				m.returnLabel = model.__global.propertyValue;
			} else {
				m.returnLabel = model.__global.returns;
			}

			newMembers.push(m);
		}
	}

	return newMembers;
}

function updateModel(model) {
	model.isTypeScript = true;

	var typeName = model.type.toLowerCase();

	if (typeName === 'class') {
		model.isClass = true;
	} else if (typeName === 'interface') {
		model.isInterface = true;
	} else if (typeName === 'type alias' || typeName === 'typealias') {
		model.isTypeAlias = true;
		typeName = 'type'; // Display '{uid} type' on type alias pages' h1 title
	} else if (typeName === 'enum') {
		model.isEnum = true;
	} else if (typeName === 'module' || typeName === 'package') {
		model.isModule = true;
	} else if (typeName === 'namespace') {
		model.isNamespace = true;
	} else if (typeName === 'container') {
		model.isContainer = true;

		if (model.landingPageType.toLowerCase() === "service") {
			model.isService = true;
		} else if (model.landingPageType.toLowerCase() === "root") {
			model.isRoot = true;
		}
	}

	if (model.name && model.name[0] && model.name[0].value) {
		var objectName = model.name[0].value;
		if (model.isContainer) {
			model.displayName = contentCommon.breakText(objectName);
		} else {
			model.displayName = contentCommon.breakText(objectName) + ' ' + typeName;
		}
	}

	if (model.deprecated) {
		if (model.deprecated.content) {
			model.deprecated.content = contentCommon.parseMarkdown(model.deprecated.content, model._key);
		}
	} else {
		model.deprecated = null;
	}

	if (!model.isPreview) {
		model.isPreview = false;
	}

	if (model.isClass) {
		if (model.children && model.children[0] && model.children[0].value && model.children[0].value.length) {
			var constructors = [],
				events = [],
				properties = [],
				methods = [];
			var children = model.children[0].value;
			var child = null,
				typeLower = null;

			for (var i = 0; i < children.length; i++) {
				child = children[i];
				typeLower = child.type.toLowerCase();

				switch (typeLower) {
					case 'constructor':
						constructors.push(child);
						break;
					case 'event':
						events.push(child);
						break;
					case 'property':
						properties.push(child);
						break;
					case 'method':
						methods.push(child);
						break;
				}
			}

			model.constructors = updateMembers(model, constructors);
			model.events = updateMembers(model, events);
			model.methods = updateMembers(model, methods);
			model.properties = updateMembers(model, properties);
		}
	} else if (model.isInterface) {
		if (model.children && model.children[0] && model.children[0].value && model.children[0].value.length) {
			var properties = [],
				methods = [];
			var children = model.children[0].value;
			var child = null;

			for (var i = 0; i < children.length; i++) {
				child = children[i];

				switch (child.type) {
					case 'property':
						properties.push(child);
						break;
					case 'method':
						methods.push(child);
						break;
				}
			}

			model.methods = updateMembers(model, methods);
			model.properties = updateMembers(model, properties);
		}
	} else if (model.isEnum) {
		if (model.children && model.children[0] && model.children[0].value && model.children[0].value.length) {
			var fields = [];
			var children = model.children[0].value;
			var child = null;

			for (var i = 0; i < children.length; i++) {
				child = children[i];

				switch (child.type) {
					case 'field':
						fields.push(child);
						break;
				}
			}

			model.fields = updateMembers(model, fields);
		}
	} else if (model.isModule || model.isNamespace) {
		if (model.children && model.children[0] && model.children[0].value && model.children[0].value.length) {
			var classes = [],
				interfaces = [],
				typeAliases = [],
				enums = [],
				properties = [],
				functions = [];
			var children = model.children[0].value;
			var child = null;

			for (var i = 0; i < children.length; i++) {
				child = children[i];

				switch (child.type) {
					case 'class':
						classes.push(child);
						break;
					case 'interface':
						interfaces.push(child);
						break;
					case 'enum':
						enums.push(child);
						break;
					case 'function':
						functions.push(child);
						break;
					case 'property':
						properties.push(child);
						break;
					case 'type alias':
					case 'typealias':
						typeAliases.push(child);
						break;
				}
			}

			model.classes = classes;
			model.interfaces = interfaces;
			model.enums = enums;
			model.functions = updateMembers(model, functions);
			model.properties = properties;
			model.typeAliases = typeAliases;
		}
	}

	if (model.inheritance && model.inheritance[0] && model.inheritance[0].value) {
		model.extendsRef = null;
		model.extends = null;
		model.inheritance[0].value.forEach(updateInheritance);
	}
	else {
		model.inheritance = null;
		model.extendsRef = null;
		if (model.extends) {
			if (Array.isArray(model.extends)) {
				// todo: Stop supporting extends as an array of string for Office TypeScript content until all content are updated.
				if (model.extends.length > 0) {
					model.extendsRef = '<xref uid=\"' + model.extends[0] + '\" displayProperty=\"name\" altProperty=\"fullName\"/>';
				}
			} else if (model.extends.href !== undefined && model.extends.name !== undefined) {
				model.extendsRef = '<a href=\"' + model.extends.href + '\">' + model.extends.name.replace('<', '&lt;').replace('>', '&gt;') + '</a>';
			} else if (model.extends.name !== undefined) {
				model.extendsRef = '<xref uid=\"' + model.extends.name.replace('<', '&lt;').replace('>', '&gt;') + '\" displayProperty=\"name\" altProperty=\"fullName\"/>';
			}
		}
	}

	model.packageRef = null;
	if (isLanguageValuePairs(model.package)) {
		model.packageRef = normalizeLanguageValuePairs(model.package).specName[0].value;
	}
	else if (model.package) {
		model.packageRef = '<xref href=\"'+ model.package + '\" />';
	}

	model.children = null;
}

function updateParameters(model, m) {
	if (!m) {
		return;
	}

	if (m.syntax && m.syntax.parameters) {
		var p = null,
			match = null;
		var newParameters = [];

		for (var j = 0; j < m.syntax.parameters.length; j++) {
			p = m.syntax.parameters[j];
			p.htmlId = m.id + "-" + contentCommon.createHtmlId(p.id);

			formatTypeStrings(model, p.type);

			newParameters.push(p);
		}

		m.syntax.parameters = newParameters;

		for (var j = 0; j < m.syntax.parameters.length; j++) {
			p = m.syntax.parameters[j];
			if (p) {
				// if (p.hasOwnProperty('optional') && p.optional) {
				// 	p.showParameterDetails = true;
				// 	p.required = (!p.optional).toString().toLowerCase();
				// }
				if (j > 0) {
					p.className = ' stack';
				}
			}
		}
	}
}

function isLanguageValuePairs(value) {
	return Array.isArray(value)
		&& value.length > 1
		&& !!value[0]
		&& !!value[0].lang
		&& !!value[0].value;
}

function normalizeLanguageValuePairs(value) {
	if (isLanguageValuePairs(value)) {
		return value[0].value;
	}
	return value;
}

function updateInheritance(tree) {
	if (!tree.type) tree.type = null;
	if (!tree.inheritance) {
		tree.inheritance = null;
	}
	else {
		tree.inheritance.forEach(updateInheritance);
	}
}

exports.updateModel = updateModel;
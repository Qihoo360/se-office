/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

$(function () {

	let logicDocument = AscTest.CreateLogicDocument();

	QUnit.module("Check text input in the document editor");

	let GetParagraphText = AscTest.GetParagraphText;

	QUnit.test("EnterText/CorrectEnterText/CompositeInput", function (assert)
	{
		AscTest.ClearDocument();
		
		let p = new AscWord.CParagraph(AscTest.DrawingDocument);
		logicDocument.AddToContent(0, p);
		
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(), "", "Check empty selection");
		
		logicDocument.AddTextWithPr("Hello World!");
		
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false, {NewLineParagraph : true}), "Hello World!\r\n", "Add text 'Hello World!'");
		
		logicDocument.MoveCursorToStartPos();
		logicDocument.MoveCursorRight();
		logicDocument.MoveCursorRight();
		
		AscTest.EnterText("123");
		assert.strictEqual(GetParagraphText(p), "He123llo World!", "Add text '123'");
		
		AscTest.EnterText("AA");
		assert.strictEqual(GetParagraphText(p), "He123AAllo World!", "Add text 'AA'");
		
		AscTest.CorrectEnterText("AB", "ABC");
		assert.strictEqual(GetParagraphText(p), "He123AAllo World!", "Check wrong correction AB to ABC");
		
		AscTest.CorrectEnterText("AA", "ABC");
		assert.strictEqual(GetParagraphText(p), "He123ABCllo World!", "Check correction AA to ABC");
		
		AscTest.EnterText("DD");
		logicDocument.MoveCursorLeft();
		AscTest.CorrectEnterText("DD", "CC");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!", "Add text DD move left and check wrong correction");
		
		logicDocument.MoveCursorToEndPos();
		AscTest.EnterText("qq");
		AscTest.CorrectEnterText("!qq", "!?");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?", "Move to the end, add qq and correct !qq to !?");
		
		AscTest.BeginCompositeInput();
		AscTest.ReplaceCompositeInput("WWW");
		AscTest.ReplaceCompositeInput("123");
		AscTest.EndCompositeInput();
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?123", "Add text '123' with composite input");
		
		AscTest.EnterTextCompositeInput("Zzz");
		AscTest.CorrectEnterText("3Zzz", "$");
		assert.strictEqual(GetParagraphText(p), "He123ABCDDllo World!?12$", "Add text 'Zzz' with composite input and correct it from '3Zzz' to '$'");
		
	});
	QUnit.test("EnterText/CorrectEnterText/CompositeInput in collaboration", function (assert)
	{
		AscTest.StartCollaboration();
		
		AscTest.ClearDocument();
		let p = new AscWord.CParagraph(AscTest.DrawingDocument);
		logicDocument.AddToContent(0, p);
		
		AscTest.EnterText("ABC");
		assert.strictEqual(GetParagraphText(p), "ABC", "Add text 'ABC' in collaboration");
		
		AscTest.MoveCursorLeft();
		AscTest.EnterText("111");
		AscTest.SyncCollaboration();
		AscTest.CorrectEnterText("11", "23");
		AscTest.SyncCollaboration();
		assert.strictEqual(GetParagraphText(p), "AB123C", "Add text '111' and correct it with '123' in collaboration (sync between actions)");
		
		AscTest.EnterText("QQQ");
		AscTest.CorrectEnterText("QQ", "RS");
		AscTest.SyncCollaboration();
		assert.strictEqual(GetParagraphText(p), "AB123QRSC", "Add text '111' and correct it with '123' in collaboration (no sync between actions)");
		
		AscTest.EndCollaboration();
	});
	QUnit.test("Test 'complex script' property on input", function (assert)
	{
		function CheckParagraphSplit(p, text, runCS, runText)
		{
			let count = runCS.length;
			assert.strictEqual(p.GetElementsCount(), count, "Check runs count on entering '" + text + "'");

			for (let index = 0; index < count; ++index)
			{
				let element = p.GetElement(index);

				if (!(element instanceof AscWord.CRun))
				{
					assert.ok(false, "Not a run");
					break;
				}

				assert.strictEqual(element.GetText(), runText[index], `Check run[${index}] text`);
				assert.strictEqual(element.IsCS(), runCS[index], `Check run[${index}] CS`);
			}
		}
		function CheckTextEnter(arrText, arrComposite, ...args)
		{
			AscTest.ClearDocument();

			let p = new AscWord.CParagraph(AscTest.DrawingDocument);
			logicDocument.AddToContent(0, p);

			let overallText = "";
			for (let index = 0, count = arrText.length; index < count; ++index)
			{
				let text = arrText[index];
				if (arrComposite[index])
					AscTest.EnterTextCompositeInput(text);
				else
					AscTest.EnterText(text);

				overallText += text;
			}

			CheckParagraphSplit(p, overallText, ...args);
		}


		CheckTextEnter(
			["Abcෑඒ"],
			[false],
			[false, true],
			["Abc", "ෑඒ"]

		);
		CheckTextEnter(
			["Abc1ෑඒ"],
			[false],
			[false, true],
			["Abc1", "ෑඒ"]
		);
		CheckTextEnter(
			["Abc1ෑඒabc"],
			[false],
			[false, true, false],
			["Abc1", "ෑඒ", "abc"]
		);

		// Композитный ввод всегда добавляет новый ран
		CheckTextEnter(
			["Abc", "ෑඒ"],
			[false, true],
			[false, true],
			["Abc", "ෑඒ"]
		);

		CheckTextEnter(
			["Abc", "ෑඒ", "efg"],
			[false, true, false],
			[false, true, false],
			["Abc", "ෑඒ", "efg"]
		);

		// Проверяем, что если ран состоит только из Script_Common, то мы параметр CS подхватывается из следующего ввода
		CheckTextEnter(
			["1abc"],
			[false],
			[false],
			["1abc"]
		);

		CheckTextEnter(
			["2ෑඒ"],
			[false],
			[true],
			["2ෑඒ"]
		);

		CheckTextEnter(
			["3", "ෑඒ"],
			[false, true],
			[true, true],
			["3", "ෑඒ"]
		);

	});
});

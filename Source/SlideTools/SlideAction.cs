using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using ActionEngine;
using DocumentFormat.OpenXml.Drawing;
using Flee.PublicTypes;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using ShapeCrawler;

using static ActionEngine.ActionEngineUtil;

namespace SlideTools
{
	//*-------------------------------------------------------------------------*
	//*	SlideActionCollection																										*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of SlideActionItem Items.
	/// </summary>
	public class SlideActionCollection :
		ActionCollectionBase<SlideActionItem, SlideActionCollection>
	{
		//*************************************************************************
		//*	Private																																*
		//*************************************************************************
		//*************************************************************************
		//*	Protected																															*
		//*************************************************************************
		//*************************************************************************
		//*	Public																																*
		//*************************************************************************

	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	SlideActionItem																													*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Information about an individual PowerPoint Control Language action.
	/// </summary>
	public class SlideActionItem :
		ActionItemBase<SlideActionItem, SlideActionCollection>
	{
		//*************************************************************************
		//*	Private																																*
		//*************************************************************************
		//*-----------------------------------------------------------------------*
		//* AlignLeft																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Align the objects in the selected objects list to the left coordinate
		/// of the last object in the list.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void AlignLeft(SlideActionItem item)
		{

		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ChangeImage																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Change the images in the selected objects list to the active working
		/// image.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void ChangeImage(SlideActionItem item)
		{

		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DistributeVertically																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Vertically distribute the objects in the selected objects list.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void DistributeVertically(SlideActionItem item)
		{
			int count = 0;
			int index = 0;
			decimal maxY = 0m;
			decimal minY = 0m;
			decimal perItem = 0m;
			List<IShape> selectedShapes = null;
			decimal y = 0m;

			if(item != null && SelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList().OrderBy(y => y.Y).ToList();
				count = selectedShapes.Count;
				if(count > 1)
				{
					minY = selectedShapes[0].Y;
					maxY = selectedShapes[^1].Y;
					perItem = (maxY - minY) / ((decimal)count - 1m);
					y = minY;
					for(index = 1; index < count; index ++)
					{
						y += perItem;
						selectedShapes[index].Y = y;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DistributeHorizontally																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Horizontally distribute the objects in the selected objects list.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void DistributeHorizontally(SlideActionItem item)
		{
			int count = 0;
			int index = 0;
			decimal maxX = 0m;
			decimal minX = 0m;
			decimal perItem = 0m;
			List<IShape> selectedShapes = null;
			decimal x = 0m;

			if(item != null && SelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList().OrderBy(x => x.X).ToList();
				count = selectedShapes.Count;
				if(count > 1)
				{
					minX = selectedShapes[0].X;
					maxX = selectedShapes[^1].X;
					perItem = (maxX - minX) / ((decimal)count - 1m);
					x = minX;
					for(index = 1; index < count; index++)
					{
						x += perItem;
						selectedShapes[index].X = x;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ForEachSelected																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Loop through each item in the selected items list, performing a set
		/// of actions on each selected item.
		/// </summary>
		/// <param name="item">
		/// Reference to the slide action item to process.
		/// </param>
		private static void ForEachSelected(SlideActionItem item)
		{

		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ForEachSlide																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Loop through each slide in the Presentation property of the local
		/// working document, running the defined child actions in the context of
		/// each slide.
		/// </summary>
		/// <param name="item">
		/// Reference to the slide action item to process.
		/// </param>
		private static async void ForEachSlide(SlideActionItem item)
		{
			int slideCount = 0;
			int slideIndex = 0;

			if(item?.WorkingDocument != null &&
				item.WorkingDocument is SlideWorkingDocumentItem workingDocument)
			{
				slideCount = workingDocument.Presentation.Slides.Count;
				for(slideIndex = 0; slideIndex < slideCount; slideIndex ++)
				{
					CurrentSlideIndex = slideIndex;
					Trace.WriteLine($"*** CurrentSlideIndex: {slideIndex} ***",
						$"{MessageImportanceEnum.Info}");
					if(item.Actions.Count > 0)
					{
						await RunActions(item.Actions);
					}
					Trace.WriteLine($"*** End Slide: {slideIndex}. Next Slide. ***",
						$"{MessageImportanceEnum.Info}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FindObjects																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Find and select all objects matching the caller-specified filter
		/// elements.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void FindObjects(SlideActionItem item)
		{
			bool bContinue = true;
			bool bMatch = false;
			string condition = "";
			bool conditionResult = false;
			ExpressionContext context = null;
			IDynamicExpression dynCondition = null;
			MatchCollection matches = null;
			string name = "";
			int slideIndex = 0;
			List<string> variableNames = null;

			if(item?.WorkingDocument != null &&
				item.WorkingDocument is SlideWorkingDocumentItem workingDocument)
			{
				if(item.Condition?.Length > 0)
				{
					condition = ResolveExpressionVariables<
						SlideActionItem, SlideActionCollection>(
							item.Condition, item);
				}
				else
				{
					condition = "(empty)";
					bContinue = false;
				}
				Trace.WriteLine(
					$" Condition: {(condition?.Length > 0 ? condition : "(empty)")}",
					$"{MessageImportanceEnum.Info}");
				if(bContinue)
				{
					variableNames = new List<string>();
					SelectedItems.Clear();
					matches = Regex.Matches(condition,
						ResourceMain.rxExpressionPairs);
					foreach(Match matchItem in matches)
					{
						name = GetValue(matchItem, "name");
						switch(name)
						{
							case "FontName":
							case "FontSize":
							case "IsBullet":
							case "SlideIndex":
								variableNames.Add(name);
								break;
							default:
								Trace.WriteLine(
									$"Error evaluating expression: {item.Condition}\r\n" +
									$"  '{GetValue(matchItem, "name")}' is not " +
									"a recognized value.");
								bContinue = false;
								break;
						}
						if(!bContinue)
						{
							break;
						}
					}
				}
				if(bContinue)
				{
					//	Variables have been assigned.
					if(variableNames.Count > 0)
					{
						context = new ExpressionContext();
						context.Imports.AddType(typeof(Math));
						context.Variables["CurrentSlideIndex"] = CurrentSlideIndex;
						slideIndex = 0;
						foreach(IUserSlide slideItem in
							workingDocument.Presentation.Slides)
						{
							if(variableNames.Contains("SlideIndex"))
							{
								context.Variables["SlideIndex"] = slideIndex;
							}
							foreach(IShape shapeItem in slideItem.Shapes)
							{
								foreach(string variableNameItem in variableNames)
								{
									switch(variableNameItem)
									{
										case "FontName":
											context.Variables["FontName"] = GetFontName(shapeItem);
											break;
										case "FontSize":
											context.Variables["FontSize"] = GetFontSize(shapeItem);
											break;
										case "IsBullet":
											context.Variables["IsBullet"] = IsBullet(shapeItem);
											break;
									}
								}
								try
								{
									dynCondition =
										context.CompileDynamic(condition.Replace('\'', '"'));
									conditionResult = (bool)dynCondition.Evaluate();
								}
								catch(Exception ex)
								{
									Trace.WriteLine(
										$"Error evaluating expression: {item.Condition}\r\n" +
										$"  {ex.Message}");
								}
								if(conditionResult)
								{
									SelectedItems.Add(shapeItem);
								}
							}
							slideIndex++;
						}
					}
					else
					{
						//	No filters. All objects are selected.
						foreach(IUserSlide slideItem in
							workingDocument.Presentation.Slides)
						{
							foreach(IShape shapeItem in slideItem.Shapes)
							{
								SelectedItems.Add(shapeItem);
							}
						}
					}
				}
				Trace.WriteLine($" Items Selected: {SelectedItems.Count}",
					$"{MessageImportanceEnum.Info}");
			}
			else
			{
				Trace.WriteLine("Error: The FindObjects action requires an " +
					"active working document.");
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetFontList																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the list of font names and sizes for the text in the provided
		/// shape.
		/// </summary>
		/// <param name="shapeItem">
		/// Reference to the shape item to inspect.
		/// </param>
		/// <returns>
		/// The list of font names and sizes for the text in the provided shape.
		/// </returns>
		private static string GetFontList(IShape shapeItem)
		{
			List<string> entries = null;
			string entry = "";
			string result = "";

			if(shapeItem?.TextBox != null)
			{
				entries = new List<string>();
				foreach(IParagraph paragraphItem in shapeItem.TextBox.Paragraphs)
				{
					foreach(IParagraphPortion portionItem in paragraphItem.Portions)
					{
						entry = $"{portionItem.Font.LatinName}:{portionItem.Font.Size}";
						if(!entries.Contains(entry))
						{
							entries.Add(entry);
						}
					}
				}
				if(entries.Count > 0)
				{
					result = string.Join(';', entries);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetFontName																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the first font name for the text in the provided shape.
		/// </summary>
		/// <param name="shapeItem">
		/// Reference to the shape item to inspect.
		/// </param>
		/// <returns>
		/// The first font name found for the text in the provided shape.
		/// </returns>
		private static string GetFontName(IShape shapeItem)
		{
			string entry = "";
			string result = "";

			if(shapeItem?.TextBox != null)
			{
				foreach(IParagraph paragraphItem in shapeItem.TextBox.Paragraphs)
				{
					foreach(IParagraphPortion portionItem in paragraphItem.Portions)
					{
						entry = $"{portionItem.Font.LatinName}";
						if(result?.Length > 0)
						{
							break;
						}
					}
					if(entry?.Length > 0)
					{
						break;
					}
				}
				if(entry?.Length > 0)
				{
					result = entry;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetFontSize																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the first font size found on the provided shape.
		/// </summary>
		/// <param name="shapeItem">
		/// Reference to the shape item to inspect.
		/// </param>
		/// <returns>
		/// The first font size found for the text in the provided shape.
		/// </returns>
		private static decimal GetFontSize(IShape shapeItem)
		{
			decimal entry = 0m;
			decimal result = 0m;

			if(shapeItem?.TextBox != null)
			{
				foreach(IParagraph paragraphItem in shapeItem.TextBox.Paragraphs)
				{
					foreach(IParagraphPortion portionItem in paragraphItem.Portions)
					{
						entry = portionItem.Font.Size;
						if(entry != 0m)
						{
							break;
						}
					}
					if(entry != 0m)
					{
						break;
					}
				}
				if(entry != 0m)
				{
					result = entry;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSelectedMaxY																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the maximum Y coordinate of the selected objects to the
		/// specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void GetSelectedMaxY(SlideActionItem item)
		{
			decimal maxY = decimal.MinValue;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList();
				if(selectedShapes.Count > 0)
				{
					maxY = selectedShapes.Max(y => y.Y);
				}
				if(maxY != decimal.MinValue)
				{
					SetVariable(item, item.VariableName, maxY);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSelectedMaxYItem																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the reference of the item having the maximum Y coordinate of all
		/// selected objects to the specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void GetSelectedMaxYItem(SlideActionItem item)
		{
			IShape maxYItem = null;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList();
				if(selectedShapes.Count > 0)
				{
					maxYItem = selectedShapes.MaxBy(y => y.Y);
				}
				SetVariable(item, item.VariableName, maxYItem);
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSelectedMinX																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the minimum X coordinate of the selected objects to the
		/// specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void GetSelectedMinX(SlideActionItem item)
		{
			decimal minX = decimal.MinValue;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList();
				if(selectedShapes.Count > 0)
				{
					minX = selectedShapes.Min(x => x.X);
				}
				if(minX != decimal.MinValue)
				{
					SetVariable(item, item.VariableName, minX);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSelectedMinY																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the minimum Y coordinate of the selected objects to the
		/// specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void GetSelectedMinY(SlideActionItem item)
		{
			decimal minY = decimal.MinValue;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList();
				if(selectedShapes.Count > 0)
				{
					minY = selectedShapes.Min(y => y.Y);
				}
				if(minY != decimal.MinValue)
				{
					SetVariable(item, item.VariableName, minY);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSelectedMinYItem																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the reference of the item having the minimum Y coordinate of all
		/// selected objects to the specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void GetSelectedMinYItem(SlideActionItem item)
		{
			IShape minYItem = null;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				selectedShapes = SelectedToShapeList();
				if(selectedShapes.Count > 0)
				{
					minYItem = selectedShapes.MinBy(y => y.Y);
				}
				SetVariable(item, item.VariableName, minYItem);
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IsBullet																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the shape has bullet-point text.
		/// </summary>
		/// <param name="shape">
		/// Reference to the shape to be inspected.
		/// </param>
		/// <returns>
		/// True if the supplied shape is a bullet point. Otherwise, false.
		/// </returns>
		private static bool IsBullet(IShape shape)
		{
			bool result = false;

			if(shape != null)
			{
				result = (shape.TextBox?.Paragraphs.FirstOrDefault(x =>
					x.Bullet?.Type != BulletType.None) != null);
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* PrintShape																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Print the contents of the provided shape, drilling into group
		/// contents where necessary.
		/// </summary>
		/// <param name="builder">
		/// Reference to the string builder to which the content is being written.
		/// </param>
		/// <param name="shapeItem">
		/// Reference to the shape item to print.
		/// </param>
		/// <param name="indent">
		/// The indent at which to print text lines.
		/// </param>
		private static void PrintShape(StringBuilder builder, IShape shapeItem,
			int indent)
		{
			bool bulleted = false;
			string filename = "";
			string fonts = "";
			IShapeCollection groupedShapes = null;
			string leader = "";
			int localIndent = indent;
			string status = "";
			string typeName = "";

			if(shapeItem != null)
			{
				leader = new string(' ', indent);
				groupedShapes = shapeItem.GroupedShapes;
				if(groupedShapes == null)
				{
					LogAppend(builder, $"{leader}{shapeItem.Name}");
					localIndent++;
					leader = new string(' ', localIndent);
					typeName = shapeItem.GetType().Name;
					if(typeName == "TextShape")
					{
						Trace.WriteLine("SlideActionItem.PrintShape: Break here...");
					}
					LogAppend(builder, $"{leader}T: {typeName}");
					LogAppend(builder, $"{leader}X: {shapeItem.X / 72m:0.##}");
					LogAppend(builder, $"{leader}Y: {shapeItem.Y / 72m:0.##}");
					LogAppend(builder, $"{leader}W: {shapeItem.Width / 72m:0.##}");
					LogAppend(builder, $"{leader}H: {shapeItem.Height / 72m:0.##}");
					switch(typeName)
					{
						case "PictureShape":
							status = shapeItem.Picture?.Image != null ? "not null" : "null";
							LogAppend(builder, $"{leader}-> {status}");
							if(status == "not null")
							{
								//filename = @"C:\Temp\" + Guid.NewGuid().ToString("D") +
								//	ActionEngineUtil.GetExtension(shapeItem.Picture.Image.Name);
								//Trace.WriteLine($"{leader}-> {filename}");
								//File.WriteAllBytes(filename,
								//	shapeItem.Picture.Image.AsByteArray());
							}
							break;
						case "TextShape":
							LogAppend(builder, $"{leader}-> {shapeItem.TextBox.Text}");
							fonts = GetFontList(shapeItem);
							LogAppend(builder, $"{leader}F: {fonts}");
							if(fonts.Length > 0)
							{
								bulleted = (shapeItem.TextBox.Paragraphs.FirstOrDefault(x =>
									x.Bullet?.Type != BulletType.None) != null);
							}
							else
							{
								bulleted = false;
							}
							LogAppend(builder,
								$"{leader}B: {(bulleted ? "" : "No ")}Bullet");
							break;
					}
				}
				else
				{
					LogAppend(builder, $"{leader}{shapeItem.Name}: GROUP");
					localIndent++;
					foreach(IShape childShapeItem in groupedShapes)
					{
						PrintShape(builder, childShapeItem, localIndent);
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SelectedToShapeList																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a specialized list of shape items from the selected items list.
		/// </summary>
		/// <returns>
		/// Reference to a list of shapes that were found in the selected items
		/// list.
		/// </returns>
		private static List<IShape> SelectedToShapeList()
		{
			List<IShape> result = new List<IShape>();

			foreach(object selectedItem in SelectedItems)
			{
				if(selectedItem is IShape shape)
				{
					result.Add(shape);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SetItemYFromVariable																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the Y coordinate of the specified item from the value in the
		/// specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void SetItemYFromVariable(SlideActionItem item)
		{
			bool bContinue = true;
			object shapeObject = null;
			decimal y = 0m;

			if(item?.VariableName?.Length > 0 && item.ItemName?.Length > 0)
			{
				shapeObject = GetVariable(item, item.ItemName);
				if(shapeObject != null)
				{
					if(shapeObject is IShape shape)
					{
						try
						{
							y = Convert.ToDecimal(GetVariable(item, item.VariableName));
						}
						catch(Exception ex)
						{
							Trace.WriteLine(
								$"Error reading variable {item.VariableName}\r\n" +
								$"  {ex.Message}",
								$"{MessageImportanceEnum.Err}");
							bContinue = false;
						}
						if(bContinue)
						{
							shape.Y = y;
						}
					}
					else
					{
						Trace.WriteLine(
							$"Error: Item is not of an IShape type: {item.ItemName}",
							$"{MessageImportanceEnum.Err}");
					}
				}
				else
				{
					Trace.WriteLine($"Error: Shape not found for {item.ItemName}",
						$"{MessageImportanceEnum.Err}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SetSelectedMaxWidth																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Set the maximum width of all items in the selection.
		/// </summary>
		/// <param name="item">
		/// Reference to the item to be limited in width.
		/// </param>
		private static void SetSelectedMaxWidth(SlideActionItem item)
		{
			bool bAdjustHeight = false;
			decimal difference = 0m;
			decimal maxValue = 0m;
			NameValueItem property = null;
			List<IShape> selectedShapes = null;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				property = item.Properties.FirstOrDefault(x =>
					StringComparer.OrdinalIgnoreCase.Equals(x.Name,
						"AdjustHeightRelative"));
				if(property != null)
				{
					bAdjustHeight = ToBool(property.Value);
				}
				try
				{
					maxValue = Convert.ToDecimal(GetVariable(item, item.VariableName));
					selectedShapes = SelectedToShapeList();
					foreach(IShape shapeItem in selectedShapes)
					{
						if(shapeItem.Width > maxValue)
						{
							if(bAdjustHeight)
							{
								difference = (shapeItem.Width - maxValue) * 0.2m;
								shapeItem.Width = maxValue;
								shapeItem.Height += difference;
							}
							else
							{
								shapeItem.Width = maxValue;
							}
						}
					}
				}
				catch(Exception ex)
				{
					Trace.WriteLine("Error converting variable to decimal: " +
						$"[{item.VariableName}] = " +
						$"{GetVariable(item, item.VariableName)}\r\n" +
						$"  {ex.Message}",
						$"{MessageImportanceEnum.Err}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SetSelectedX																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the X coordinate on all selected objects to the value in the
		/// specified variable.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void SetSelectedX(SlideActionItem item)
		{
			bool bContinue = true;
			List<IShape> selectedShapes = null;
			decimal x = 0m;

			if(item?.VariableName?.Length > 0 && mSelectedItems.Count > 0)
			{
				try
				{
					x = Convert.ToDecimal(GetVariable(item, item.VariableName));
				}
				catch(Exception ex)
				{
					Trace.WriteLine($"Error reading variable {item.VariableName}\r\n" +
						$"  {ex.Message}",
						$"{MessageImportanceEnum.Err}");
					bContinue = false;
				}
				if(bContinue)
				{
					selectedShapes = SelectedToShapeList();
					foreach(IShape shapeItem in selectedShapes)
					{
						shapeItem.X = x;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SlideReport																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Create a report of all the slides in the specified file.
		/// </summary>
		/// <param name="item">
		/// Reference to the active action.
		/// </param>
		private static void SlideReport(SlideActionItem item)
		{
			StringBuilder builder = null;
			int indent = 0;
			int index = 0;
			Presentation presentation = null;

			if(item != null)
			{
				if(CheckElements(item, ActionElementEnum.InputFilename))
				{
					builder = new StringBuilder();

					LogAppend(builder, "*** SLIDE REPORT ***");
					presentation = new Presentation(item.InputFiles[0].FullName);
					if(presentation.Slides.Count > 0)
					{
						LogAppend(builder,
							$"Width:       {presentation.SlideWidth / 72m:0.##}");
						LogAppend(builder,
							$"Height:      {presentation.SlideHeight / 72m:0.##}");
						LogAppend(builder, "");
						LogAppend(builder, $"Slide Count: {presentation.Slides.Count}");
						index = 1;
						indent = 1;
						foreach(IUserSlide userSlideItem in presentation.Slides)
						{
							LogAppend(builder, "");
							LogAppend(builder, $"Slide {index}");
							LogAppend(builder, new string('-', 40));
							foreach(IShape shapeItem in userSlideItem.Shapes)
							{
								PrintShape(builder, shapeItem, indent);
							}
							LogAppend(builder, new string('-', 40));
							index++;
						}
					}
					presentation.Dispose();
					presentation = null;
					if(item.OutputFile != null)
					{
						File.WriteAllText(item.OutputFile.FullName, builder.ToString());
						Trace.WriteLine(
							$"Report output written to: {item.OutputFile.Name}");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*************************************************************************
		//*	Protected																															*
		//*************************************************************************
		////*-----------------------------------------------------------------------*
		////* DeserializeFile																												*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Deserialize the this specialized version of the object model using
		///// the caller's supplied content.
		///// </summary>
		///// <param name="content">
		///// The JSON-formatted content representing the specialized data fitting
		///// to this class.
		///// </param>
		///// <returns>
		///// Reference to a PAction item containing the deserialized object model,
		///// if legitimate. Otherwise, null.
		///// </returns>
		//protected override SlideActionItem DeserializeFile(string content)
		//{
		//	SlideActionItem result = null;

		//	if(content?.Length > 0)
		//	{
		//		try
		//		{
		//			result = JsonConvert.DeserializeObject<SlideActionItem>(content);
		//		}
		//		catch(Exception ex)
		//		{
		//			Trace.WriteLine($"Error deserializing: {ex.Message}",
		//				MessageImportanceEnum.Err.ToString());
		//		}
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeExpressionValues																						*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize specialized expression values from the implemented instance.
		/// </summary>
		/// <param name="context">
		/// Reference to the expression context to which any additions or changes
		/// will be made.
		/// </param>
		protected override void InitializeExpressionValues(
			ExpressionContext context)
		{
			base.InitializeExpressionValues(context);
			context.Variables["SelectedItemCount"] = SelectedItems.Count;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeVariable																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize and return the supplied variable under preparation.
		/// </summary>
		/// <param name="variable">
		/// Reference to the variable being prepared.
		/// </param>
		/// <remarks>
		/// If the variable initialization is cancelled, its name will be set to
		/// an empty string and its value will be set to null.
		/// </remarks>
		protected override void InitializeVariable(VariableItem variable)
		{
			bool bUpdate = false;
			decimal binary = 0m;
			Match match = null;
			string number = "";
			string unit = "";

			if(variable?.Name?.Length > 0 && variable.Value is string varValue)
			{
				match = Regex.Match(varValue, ResourceMain.rxNumberWithUnit);
				if(match.Success)
				{
					number = GetValue(match, "number");
					unit = GetValue(match, "unit").ToLower();
					bUpdate = true;
					switch(unit)
					{
						case "d":
							binary = Convert.ToDecimal(number);
							break;
						case "in":
							binary = Convert.ToDecimal(number) * 72m;
							break;
						case "mm":
							binary = (Convert.ToDecimal(number) / 25.4m) * 72m;
							break;
						default:
							bUpdate = false;
							break;
					}
					if(bUpdate)
					{
						variable.Value = binary;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* OpenWorkingDocument																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Open the working file to allow multiple operations to be completed
		/// in the same session.
		/// </summary>
		/// <param name="item">
		/// Reference to the action item containing information about the file to
		/// open.
		/// </param>
		protected override void OpenWorkingDocument()
		{
			string content = "";
			ActionDocumentItem doc = null;
			int docIndex = 0;

			if(CheckElements(this,
				ActionElementEnum.InputFilename))
			{
				//	Load the document if a filename was specified.
				docIndex = WorkingDocumentIndex;
				if(docIndex > -1 && docIndex < InputFiles.Count)
				{
					WorkingDocument =
						new SlideWorkingDocumentItem(InputFiles[docIndex].FullName);
					Trace.WriteLine(
						$" Working document: {this.InputFiles[docIndex].Name}",
						$"{MessageImportanceEnum.Info}");
				}
				else
				{
					Trace.WriteLine(
						$" Working document index out of range at: {docIndex}",
						$"{MessageImportanceEnum.Warn}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RunCustomAction																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the custom action.
		/// </summary>
		protected override void RunCustomAction()
		{
			string action = Action.ToLower();

			switch(action)
			{
				case "alignleft":
					AlignLeft(this);
					break;
				case "changeimage":
					ChangeImage(this);
					break;
				case "distributehorizontally":
					DistributeHorizontally(this);
					break;
				case "distributevertically":
					DistributeVertically(this);
					break;
				case "findobjects":
					FindObjects(this);
					break;
				case "foreachselected":
					ForEachSelected(this);
					break;
				case "foreachslide":
					ForEachSlide(this);
					break;
				case "getselectedmaxy":
					GetSelectedMaxY(this);
					break;
				case "getselectedmaxyitem":
					GetSelectedMaxYItem(this);
					break;
				case "getselectedminx":
					GetSelectedMinX(this);
					break;
				case "getselectedminy":
					GetSelectedMinY(this);
					break;
				case "getselectedminyitem":
					GetSelectedMinYItem(this);
					break;
				case "setitemyfromvariable":
					SetItemYFromVariable(this);
					break;
				case "setselectedmaxwidth":
					SetSelectedMaxWidth(this);
					break;
				case "setselectedx":
					SetSelectedX(this);
					break;
				case "slidereport":
					SlideReport(this);
					break;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SaveWorkingDocument																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Save the working file to the specified output file.
		/// </summary>
		protected override void SaveWorkingDocument()
		{
			if(WorkingDocument != null &&
				WorkingDocument is SlideWorkingDocumentItem workingDocument &&
				CheckElements(this, ActionElementEnum.OutputFilename))
			{
				//	Document is open and output file has been specified.
				try
				{
					using(MemoryStream memoryStream = new MemoryStream())
					{
						workingDocument.Presentation.Save(memoryStream);
						using(FileStream fileStream = File.Create(OutputFile.FullName))
						{
							memoryStream.Position = 0;
							memoryStream.CopyTo(fileStream);
						}
					}
					Trace.WriteLine($" Document saved to: {OutputFile.Name}",
						$"{MessageImportanceEnum.Info}");
				}
				catch(Exception ex)
				{
					Trace.WriteLine(
						$"Error while saving document: {OutputFile.Name}\r\n" +
						$"  {ex.Message}",
						$"{MessageImportanceEnum.Err}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* WriteLocalOutput																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Write the local output of this operation.
		/// </summary>
		protected override void WriteLocalOutput()
		{
			base.WriteLocalOutput();
		}
		//*-----------------------------------------------------------------------*

		//*************************************************************************
		//*	Public																																*
		//*************************************************************************

		////*-----------------------------------------------------------------------*
		////*	Actions																																*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Private member for <see cref="Actions">Actions</see>.
		///// </summary>
		//private SlideActionCollection mActions =
		//	new SlideActionCollection();
		///// <summary>
		///// Get a reference to the collection of child SVG actions.
		///// </summary>
		///// <remarks>
		///// This property is non-inheritable.
		///// </remarks>
		//public new SlideActionCollection Actions
		//{
		//	get { return mActions; }
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	CurrentSlideIndex																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for
		/// <see cref="CurrentSlideIndex">CurrentSlideIndex</see>.
		/// </summary>
		private static int mCurrentSlideIndex = 0;
		/// <summary>
		/// Get/Set the current slide index for this session.
		/// </summary>
		public static int CurrentSlideIndex
		{
			get { return mCurrentSlideIndex; }
			set { mCurrentSlideIndex = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Filters																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Filters">Filters</see>.
		/// </summary>
		private FilterElementCollection mFilters = new FilterElementCollection();
		/// <summary>
		/// Get a reference to the collection of filters on this action. This value
		/// is inheritable.
		/// </summary>
		public FilterElementCollection Filters
		{
			get
			{
				FilterElementCollection filters = mFilters;

				if(filters.Count == 0 && Parent?.Parent != null)
				{
					filters = ((SlideActionItem)Parent.Parent).Filters;
				}
				return filters;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ItemName																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="ItemName">ItemName</see>.
		/// </summary>
		private string mItemName = null;
		/// <summary>
		/// Get/Set the item name for this action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public string ItemName
		{
			get
			{
				string result = mItemName;

				if(result == null)
				{
					if(Parent?.Parent != null)
					{
						result = Parent.Parent.ItemName;
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mItemName = value; }
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////*	Parent																																*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Private member for <see cref="Parent">Parent</see>.
		///// </summary>
		//private SlideActionCollection mParent = null;
		///// <summary>
		///// Get/Set a reference to the parent of this item.
		///// </summary>
		///// <remarks>
		///// This property is non-inheritable.
		///// </remarks>
		//[JsonIgnore]
		//public new SlideActionCollection Parent
		//{
		//	get { return mParent; }
		//	set { mParent = value; }
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	SelectedItems																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="SelectedItems">SelectedItems</see>.
		/// </summary>
		private static List<object> mSelectedItems = new List<object>();
		/// <summary>
		/// Get a reference to the collection of items currently selected.
		/// </summary>
		public static List<object> SelectedItems
		{
			get { return mSelectedItems; }
		}
		//*-----------------------------------------------------------------------*


	}
	//*-------------------------------------------------------------------------*

}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Novacode
{

	#region Enums

	#region Internal

	/// <summary>Custom property types</summary>
	internal enum CustomPropertyType
	{

		/// <summary>System.String</summary>
		Text,

		/// <summary>System.DateTime</summary>
		Date,

		/// <summary>System.Int32</summary>
		NumberInteger,

		/// <summary>System.Double</summary>
		NumberDecimal,

		/// <summary>System.Boolean</summary>
		YesOrNo

	}

	/// <summary>Paragraph edit types</summary>
	internal enum EditType
	{

		/// <summary>A ins is a tracked insertion</summary>
		ins,

		/// <summary>A del is tracked deletion</summary>
		del

	}

	#endregion

	#region Public

	/// <summary>Specifies the possible directions for a bar chart. 21.2.3.3 ST_BarDir (Bar Direction)</summary>
	public enum BarDirection
	{

		/// <summary>TODO: comment</summary>
		[XmlName("col")]
		Column,

		/// <summary>TODO: comment</summary>
		[XmlName("bar")]
		Bar

	}

	/// <summary>Specifies the possible groupings for a bar chart. 21.2.3.4 ST_BarGrouping (Bar Grouping)</summary>
	public enum BarGrouping
	{

		/// <summary>TODO: comment</summary>
		[XmlName("clustered")]
		Clustered,

		/// <summary>TODO: comment</summary>
		[XmlName("percentStacked")]
		PercentStacked,

		/// <summary>TODO: comment</summary>
		[XmlName("stacked")]
		Stacked,

		/// <summary>TODO: comment</summary>
		[XmlName("standard")]
		Standard

	}

	/// <summary>Specifies the possible positions for a legend. 21.2.3.24 ST_LegendPos (Legend Position)</summary>
	public enum ChartLegendPosition
	{

		/// <summary>TODO: comment</summary>
		[XmlName("t")]
		Top,

		/// <summary>TODO: comment</summary>
		[XmlName("b")]
		Bottom,

		/// <summary>TODO: comment</summary>
		[XmlName("l")]
		Left,

		/// <summary>TODO: comment</summary>
		[XmlName("r")]
		Right,

		/// <summary>TODO: comment</summary>
		[XmlName("tr")]
		TopRight

	}

	/// <summary>Specifies the possible ways to display blanks. 21.2.3.10 ST_DispBlanksAs (Display Blanks As)</summary>
	public enum DisplayBlanksAs
	{

		/// <summary>TODO: comment</summary>
		[XmlName("gap")]
		Gap,

		/// <summary>TODO: comment</summary>
		[XmlName("span")]
		Span,

		/// <summary>TODO: comment</summary>
		[XmlName("zero")]
		Zero

	}

	/// <summary>Specifies the kind of grouping for a column, line, or area chart. 21.2.2.76 grouping (Grouping)</summary>
	public enum Grouping
	{

		/// <summary>TODO: comment</summary>
		[XmlName("percentStacked")]
		PercentStacked,

		/// <summary>TODO: comment</summary>
		[XmlName("stacked")]
		Stacked,

		/// <summary>TODO: comment</summary>
		[XmlName("standard")]
		Standard

	}

	/// <summary></summary>
	public enum DocumentTypes
	{

		/// <summary>TODO: comment</summary>
		Document,

		/// <summary>TODO: comment</summary>
		Template

	}

	/// <summary>TODO: comment</summary>
	public enum ListItemType
	{

		/// <summary>TODO: comment</summary>
		Bulleted,

		/// <summary>TODO: comment</summary>
		Numbered

	}

	/// <summary>TODO: comment</summary>
	public enum SectionBreakType
	{

		/// <summary>TODO: comment</summary>
		defaultNextPage,

		/// <summary>TODO: comment</summary>
		evenPage,

		/// <summary>TODO: comment</summary>
		oddPage,

		/// <summary>TODO: comment</summary>
		continuous

	}

	/// <summary>TODO: comment</summary>
	public enum ContainerType
	{

		/// <summary>TODO: comment</summary>
		None,

		/// <summary>TODO: comment</summary>
		TOC,

		/// <summary>TODO: comment</summary>
		Section,

		/// <summary>TODO: comment</summary>
		Cell,

		/// <summary>TODO: comment</summary>
		Table,

		/// <summary>TODO: comment</summary>
		Header,

		/// <summary>TODO: comment</summary>
		Footer,

		/// <summary>TODO: comment</summary>
		Paragraph,

		/// <summary>TODO: comment</summary>
		Body

	}

	/// <summary>TODO: comment</summary>
	public enum PageNumberFormat
	{

		/// <summary>TODO: comment</summary>
		normal,

		/// <summary>TODO: comment</summary>
		roman

	}

	/// <summary>TODO: comment</summary>
	public enum BorderSize
	{

		/// <summary>TODO: comment</summary>
		one,

		/// <summary>TODO: comment</summary>
		two,

		/// <summary>TODO: comment</summary>
		three,

		/// <summary>TODO: comment</summary>
		four,

		/// <summary>TODO: comment</summary>
		five,

		/// <summary>TODO: comment</summary>
		six,

		/// <summary>TODO: comment</summary>
		seven,

		/// <summary>TODO: comment</summary>
		eight,

		/// <summary>TODO: comment</summary>
		nine

	}

	/// <summary>TODO: comment</summary>
	public enum EditRestrictions
	{

		/// <summary>TODO: comment</summary>
		none,

		/// <summary>TODO: comment</summary>
		readOnly,

		/// <summary>TODO: comment</summary>
		forms,

		/// <summary>TODO: comment</summary>
		comments,

		/// <summary>TODO: comment</summary>
		trackedChanges

	}

	/// <summary>Table Cell Border styles. Added by lckuiper @ 20101117. Source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx</summary>
	public enum BorderStyle
	{

		/// <summary>TODO: comment</summary>
		Tcbs_none = 0,

		/// <summary>TODO: comment</summary>
		Tcbs_single,

		/// <summary>TODO: comment</summary>
		Tcbs_thick,

		/// <summary>TODO: comment</summary>
		Tcbs_double,

		/// <summary>TODO: comment</summary>
		Tcbs_dotted,

		/// <summary>TODO: comment</summary>
		Tcbs_dashed,

		/// <summary>TODO: comment</summary>
		Tcbs_dotDash,

		/// <summary>TODO: comment</summary>
		Tcbs_dotDotDash,

		/// <summary>TODO: comment</summary>
		Tcbs_triple,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickSmallGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thickThinSmallGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickThinSmallGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickMediumGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thickThinMediumGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickThinMediumGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickLargeGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thickThinLargeGap,

		/// <summary>TODO: comment</summary>
		Tcbs_thinThickThinLargeGap,

		/// <summary>TODO: comment</summary>
		Tcbs_wave,

		/// <summary>TODO: comment</summary>
		Tcbs_doubleWave,

		/// <summary>TODO: comment</summary>
		Tcbs_dashSmallGap,

		/// <summary>TODO: comment</summary>
		Tcbs_dashDotStroked,

		/// <summary>TODO: comment</summary>
		Tcbs_threeDEmboss,

		/// <summary>TODO: comment</summary>
		Tcbs_threeDEngrave,

		/// <summary>TODO: comment</summary>
		Tcbs_outset,

		/// <summary>TODO: comment</summary>
		Tcbs_inset,

		/// <summary>TODO: comment</summary>
		Tcbs_nil

	}

	/// <summary>Table Cell Border Types. Added by lckuiper @ 20101117. Source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx</summary>
	public enum TableCellBorderType
	{

		/// <summary>TODO: comment</summary>
		Top,

		/// <summary>TODO: comment</summary>
		Bottom,

		/// <summary>TODO: comment</summary>
		Left,

		/// <summary>TODO: comment</summary>
		Right,

		/// <summary>TODO: comment</summary>
		InsideH,

		/// <summary>TODO: comment</summary>
		InsideV,

		/// <summary>TODO: comment</summary>
		TopLeftToBottomRight,

		/// <summary>TODO: comment</summary>
		TopRightToBottomLeft

	}

	/// <summary>Table Border Types. Added by lckuiper @ 20101117. Source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tableborders.aspx</summary>
	public enum TableBorderType
	{

		/// <summary>TODO: comment</summary>
		Top,

		/// <summary>TODO: comment</summary>
		Bottom,

		/// <summary>TODO: comment</summary>
		Left,

		/// <summary>TODO: comment</summary>
		Right,

		/// <summary>TODO: comment</summary>
		InsideH,

		/// <summary>TODO: comment</summary>
		InsideV

	}

	/// <summary>Patch 7398 added by lckuiper on Nov 16th 2010 @ 2:23 PM</summary>
	public enum VerticalAlignment
	{

		/// <summary>TODO: comment</summary>
		Top,

		/// <summary>TODO: comment</summary>
		Center,

		/// <summary>TODO: comment</summary>
		Bottom

	};

	/// <summary>TODO: comment</summary>
	public enum Orientation
	{

		/// <summary>TODO: comment</summary>
		Portrait,

		/// <summary>TODO: comment</summary>
		Landscape

	};

	/// <summary>TODO: comment</summary>
	public enum XmlDocument
	{

		/// <summary>TODO: comment</summary>
		Main,

		/// <summary>TODO: comment</summary>
		HeaderOdd,

		/// <summary>TODO: comment</summary>
		HeaderEven,

		/// <summary>TODO: comment</summary>
		HeaderFirst,

		/// <summary>TODO: comment</summary>
		FooterOdd,

		/// <summary>TODO: comment</summary>
		FooterEven,

		/// <summary>TODO: comment</summary>
		FooterFirst

	};

	/// <summary>TODO: comment</summary>
	public enum MatchFormattingOptions
	{

		/// <summary>TODO: comment</summary>
		ExactMatch,

		/// <summary>TODO: comment</summary>
		SubsetMatch

	};

	/// <summary>TODO: comment</summary>
	public enum Script
	{

		/// <summary>TODO: comment</summary>
		superscript,

		/// <summary>TODO: comment</summary>
		subscript,

		/// <summary>TODO: comment</summary>
		none

	}

	/// <summary>TODO: comment</summary>
	public enum Highlight
	{

		/// <summary>TODO: comment</summary>
		yellow,

		/// <summary>TODO: comment</summary>
		green,

		/// <summary>TODO: comment</summary>
		cyan,

		/// <summary>TODO: comment</summary>
		magenta,

		/// <summary>TODO: comment</summary>
		blue,

		/// <summary>TODO: comment</summary>
		red,

		/// <summary>TODO: comment</summary>
		darkBlue,

		/// <summary>TODO: comment</summary>
		darkCyan,

		/// <summary>TODO: comment</summary>
		darkGreen,

		/// <summary>TODO: comment</summary>
		darkMagenta,

		/// <summary>TODO: comment</summary>
		darkRed,

		/// <summary>TODO: comment</summary>
		darkYellow,

		/// <summary>TODO: comment</summary>
		darkGray,

		/// <summary>TODO: comment</summary>
		lightGray,

		/// <summary>TODO: comment</summary>
		black,

		/// <summary>TODO: comment</summary>
		none

	};

	/// <summary>TODO: comment</summary>
	public enum UnderlineStyle
	{

		/// <summary>TODO: comment</summary>
		none = 0,

		/// <summary>TODO: comment</summary>
		singleLine = 1,

		/// <summary>TODO: comment</summary>
		words = 2,

		/// <summary>TODO: comment</summary>
		doubleLine = 3,

		/// <summary>TODO: comment</summary>
		dotted = 4,

		/// <summary>TODO: comment</summary>
		thick = 6,

		/// <summary>TODO: comment</summary>
		dash = 7,

		/// <summary>TODO: comment</summary>
		dotDash = 9,

		/// <summary>TODO: comment</summary>
		dotDotDash = 10,

		/// <summary>TODO: comment</summary>
		wave = 11,

		/// <summary>TODO: comment</summary>
		dottedHeavy = 20,

		/// <summary>TODO: comment</summary>
		dashedHeavy = 23,

		/// <summary>TODO: comment</summary>
		dashDotHeavy = 25,

		/// <summary>TODO: comment</summary>
		dashDotDotHeavy = 26,

		/// <summary>TODO: comment</summary>
		dashLongHeavy = 27,

		/// <summary>TODO: comment</summary>
		dashLong = 39,

		/// <summary>TODO: comment</summary>
		wavyDouble = 43,

		/// <summary>TODO: comment</summary>
		wavyHeavy = 55

	};

	/// <summary>TODO: comment</summary>
	public enum StrikeThrough
	{

		/// <summary>TODO: comment</summary>
		none,

		/// <summary>TODO: comment</summary>
		strike,

		/// <summary>TODO: comment</summary>
		doubleStrike

	};

	/// <summary>TODO: comment</summary>
	public enum Misc
	{

		/// <summary>TODO: comment</summary>
		none,

		/// <summary>TODO: comment</summary>
		shadow,

		/// <summary>TODO: comment</summary>
		outline,

		/// <summary>TODO: comment</summary>
		outlineShadow,

		/// <summary>TODO: comment</summary>
		emboss,

		/// <summary>TODO: comment</summary>
		engrave

	};

	/// <summary>Change the caps style of text, for use with Append and AppendLine</summary>
	public enum CapsStyle
	{

		/// <summary>No caps, make all characters are lowercase</summary>
		none,

		/// <summary>All caps, make every character uppercase</summary>
		caps,

		/// <summary>Small caps, make all characters capital but with a small font size</summary>
		smallCaps

	};

	/// <summary>Designs\Styles that can be applied to a table</summary>
	public enum TableDesign
	{

		/// <summary>TODO: comment</summary>
		Custom,

		/// <summary>TODO: comment</summary>
		TableNormal,

		/// <summary>TODO: comment</summary>
		TableGrid,

		/// <summary>TODO: comment</summary>
		LightShading,

		/// <summary>TODO: comment</summary>
		LightShadingAccent1,

		/// <summary>TODO: comment</summary>
		LightShadingAccent2,

		/// <summary>TODO: comment</summary>
		LightShadingAccent3,

		/// <summary>TODO: comment</summary>
		LightShadingAccent4,

		/// <summary>TODO: comment</summary>
		LightShadingAccent5,

		/// <summary>TODO: comment</summary>
		LightShadingAccent6,

		/// <summary>TODO: comment</summary>
		LightList,

		/// <summary>TODO: comment</summary>
		LightListAccent1,

		/// <summary>TODO: comment</summary>
		LightListAccent2,

		/// <summary>TODO: comment</summary>
		LightListAccent3,

		/// <summary>TODO: comment</summary>
		LightListAccent4,

		/// <summary>TODO: comment</summary>
		LightListAccent5,

		/// <summary>TODO: comment</summary>
		LightListAccent6,

		/// <summary>TODO: comment</summary>
		LightGrid,

		/// <summary>TODO: comment</summary>
		LightGridAccent1,

		/// <summary>TODO: comment</summary>
		LightGridAccent2,

		/// <summary>TODO: comment</summary>
		LightGridAccent3,

		/// <summary>TODO: comment</summary>
		LightGridAccent4,

		/// <summary>TODO: comment</summary>
		LightGridAccent5,

		/// <summary>TODO: comment</summary>
		LightGridAccent6,

		/// <summary>TODO: comment</summary>
		MediumShading1,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent1,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent2,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent3,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent4,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent5,

		/// <summary>TODO: comment</summary>
		MediumShading1Accent6,

		/// <summary>TODO: comment</summary>
		MediumShading2,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent1,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent2,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent3,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent4,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent5,

		/// <summary>TODO: comment</summary>
		MediumShading2Accent6,

		/// <summary>TODO: comment</summary>
		MediumList1,

		/// <summary>TODO: comment</summary>
		MediumList1Accent1,

		/// <summary>TODO: comment</summary>
		MediumList1Accent2,

		/// <summary>TODO: comment</summary>
		MediumList1Accent3,

		/// <summary>TODO: comment</summary>
		MediumList1Accent4,

		/// <summary>TODO: comment</summary>
		MediumList1Accent5,

		/// <summary>TODO: comment</summary>
		MediumList1Accent6,

		/// <summary>TODO: comment</summary>
		MediumList2,

		/// <summary>TODO: comment</summary>
		MediumList2Accent1,

		/// <summary>TODO: comment</summary>
		MediumList2Accent2,

		/// <summary>TODO: comment</summary>
		MediumList2Accent3,

		/// <summary>TODO: comment</summary>
		MediumList2Accent4,

		/// <summary>TODO: comment</summary>
		MediumList2Accent5,

		/// <summary>TODO: comment</summary>
		MediumList2Accent6,

		/// <summary>TODO: comment</summary>
		MediumGrid1,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent1,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent2,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent3,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent4,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent5,

		/// <summary>TODO: comment</summary>
		MediumGrid1Accent6,

		/// <summary>TODO: comment</summary>
		MediumGrid2,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent1,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent2,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent3,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent4,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent5,

		/// <summary>TODO: comment</summary>
		MediumGrid2Accent6,

		/// <summary>TODO: comment</summary>
		MediumGrid3,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent1,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent2,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent3,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent4,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent5,

		/// <summary>TODO: comment</summary>
		MediumGrid3Accent6,

		/// <summary>TODO: comment</summary>
		DarkList,

		/// <summary>TODO: comment</summary>
		DarkListAccent1,

		/// <summary>TODO: comment</summary>
		DarkListAccent2,

		/// <summary>TODO: comment</summary>
		DarkListAccent3,

		/// <summary>TODO: comment</summary>
		DarkListAccent4,

		/// <summary>TODO: comment</summary>
		DarkListAccent5,

		/// <summary>TODO: comment</summary>
		DarkListAccent6,

		/// <summary>TODO: comment</summary>
		ColorfulShading,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent1,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent2,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent3,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent4,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent5,

		/// <summary>TODO: comment</summary>
		ColorfulShadingAccent6,

		/// <summary>TODO: comment</summary>
		ColorfulList,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent1,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent2,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent3,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent4,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent5,

		/// <summary>TODO: comment</summary>
		ColorfulListAccent6,

		/// <summary>TODO: comment</summary>
		ColorfulGrid,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent1,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent2,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent3,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent4,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent5,

		/// <summary>TODO: comment</summary>
		ColorfulGridAccent6,

		/// <summary>TODO: comment</summary>
		None

	};

	/// <summary>How a Table should auto resize</summary>
	public enum AutoFit
	{

		/// <summary>Autofit to Table contents</summary>
		Contents,

		/// <summary>Autofit to Window</summary>
		Window,

		/// <summary>Autofit to Column width</summary>
		ColumnWidth,

		/// <summary>Autofit to Fixed column width</summary>
		Fixed

	};

	/// <summary>TODO: comment</summary>
	public enum RectangleShapes
	{

		/// <summary>TODO: comment</summary>
		rect,

		/// <summary>TODO: comment</summary>
		roundRect,

		/// <summary>TODO: comment</summary>
		snip1Rect,

		/// <summary>TODO: comment</summary>
		snip2SameRect,

		/// <summary>TODO: comment</summary>
		snip2DiagRect,

		/// <summary>TODO: comment</summary>
		snipRoundRect,

		/// <summary>TODO: comment</summary>
		round1Rect,

		/// <summary>TODO: comment</summary>
		round2SameRect,

		/// <summary>TODO: comment</summary>
		round2DiagRect

	};

	/// <summary>TODO: comment</summary>
	public enum BasicShapes
	{

		/// <summary>TODO: comment</summary>
		ellipse,

		/// <summary>TODO: comment</summary>
		triangle,

		/// <summary>TODO: comment</summary>
		rtTriangle,

		/// <summary>TODO: comment</summary>
		parallelogram,

		/// <summary>TODO: comment</summary>
		trapezoid,

		/// <summary>TODO: comment</summary>
		diamond,

		/// <summary>TODO: comment</summary>
		pentagon,

		/// <summary>TODO: comment</summary>
		hexagon,

		/// <summary>TODO: comment</summary>
		heptagon,

		/// <summary>TODO: comment</summary>
		octagon,

		/// <summary>TODO: comment</summary>
		decagon,

		/// <summary>TODO: comment</summary>
		dodecagon,

		/// <summary>TODO: comment</summary>
		pie,

		/// <summary>TODO: comment</summary>
		chord,

		/// <summary>TODO: comment</summary>
		teardrop,

		/// <summary>TODO: comment</summary>
		frame,

		/// <summary>TODO: comment</summary>
		halfFrame,

		/// <summary>TODO: comment</summary>
		corner,

		/// <summary>TODO: comment</summary>
		diagStripe,

		/// <summary>TODO: comment</summary>
		plus,

		/// <summary>TODO: comment</summary>
		plaque,

		/// <summary>TODO: comment</summary>
		can,

		/// <summary>TODO: comment</summary>
		cube,

		/// <summary>TODO: comment</summary>
		bevel,

		/// <summary>TODO: comment</summary>
		donut,

		/// <summary>TODO: comment</summary>
		noSmoking,

		/// <summary>TODO: comment</summary>
		blockArc,

		/// <summary>TODO: comment</summary>
		foldedCorner,

		/// <summary>TODO: comment</summary>
		smileyFace,

		/// <summary>TODO: comment</summary>
		heart,

		/// <summary>TODO: comment</summary>
		lightningBolt,

		/// <summary>TODO: comment</summary>
		sun,

		/// <summary>TODO: comment</summary>
		moon,

		/// <summary>TODO: comment</summary>
		cloud,

		/// <summary>TODO: comment</summary>
		arc,

		/// <summary>TODO: comment</summary>
		backetPair,

		/// <summary>TODO: comment</summary>
		bracePair,

		/// <summary>TODO: comment</summary>
		leftBracket,

		/// <summary>TODO: comment</summary>
		rightBracket,

		/// <summary>TODO: comment</summary>
		leftBrace,

		/// <summary>TODO: comment</summary>
		rightBrace

	};

	/// <summary>TODO: comment</summary>
	public enum BlockArrowShapes
	{

		/// <summary>TODO: comment</summary>
		rightArrow,

		/// <summary>TODO: comment</summary>
		leftArrow,

		/// <summary>TODO: comment</summary>
		upArrow,

		/// <summary>TODO: comment</summary>
		downArrow,

		/// <summary>TODO: comment</summary>
		leftRightArrow,

		/// <summary>TODO: comment</summary>
		upDownArrow,

		/// <summary>TODO: comment</summary>
		quadArrow,

		/// <summary>TODO: comment</summary>
		leftRightUpArrow,

		/// <summary>TODO: comment</summary>
		bentArrow,

		/// <summary>TODO: comment</summary>
		uturnArrow,

		/// <summary>TODO: comment</summary>
		leftUpArrow,

		/// <summary>TODO: comment</summary>
		bentUpArrow,

		/// <summary>TODO: comment</summary>
		curvedRightArrow,

		/// <summary>TODO: comment</summary>
		curvedLeftArrow,

		/// <summary>TODO: comment</summary>
		curvedUpArrow,

		/// <summary>TODO: comment</summary>
		curvedDownArrow,

		/// <summary>TODO: comment</summary>
		stripedRightArrow,

		/// <summary>TODO: comment</summary>
		notchedRightArrow,

		/// <summary>TODO: comment</summary>
		homePlate,

		/// <summary>TODO: comment</summary>
		chevron,

		/// <summary>TODO: comment</summary>
		rightArrowCallout,

		/// <summary>TODO: comment</summary>
		downArrowCallout,

		/// <summary>TODO: comment</summary>
		leftArrowCallout,

		/// <summary>TODO: comment</summary>
		upArrowCallout,

		/// <summary>TODO: comment</summary>
		leftRightArrowCallout,

		/// <summary>TODO: comment</summary>
		quadArrowCallout,

		/// <summary>TODO: comment</summary>
		circularArrow

	};

	/// <summary>TODO: comment</summary>
	public enum EquationShapes
	{

		/// <summary>TODO: comment</summary>
		mathPlus,

		/// <summary>TODO: comment</summary>
		mathMinus,

		/// <summary>TODO: comment</summary>
		mathMultiply,

		/// <summary>TODO: comment</summary>
		mathDivide,

		/// <summary>TODO: comment</summary>
		mathEqual,

		/// <summary>TODO: comment</summary>
		mathNotEqual

	};

	/// <summary>TODO: comment</summary>
	public enum FlowchartShapes
	{

		/// <summary>TODO: comment</summary>
		flowChartProcess,

		/// <summary>TODO: comment</summary>
		flowChartAlternateProcess,

		/// <summary>TODO: comment</summary>
		flowChartDecision,

		/// <summary>TODO: comment</summary>
		flowChartInputOutput,

		/// <summary>TODO: comment</summary>
		flowChartPredefinedProcess,

		/// <summary>TODO: comment</summary>
		flowChartInternalStorage,

		/// <summary>TODO: comment</summary>
		flowChartDocument,

		/// <summary>TODO: comment</summary>
		flowChartMultidocument,

		/// <summary>TODO: comment</summary>
		flowChartTerminator,

		/// <summary>TODO: comment</summary>
		flowChartPreparation,

		/// <summary>TODO: comment</summary>
		flowChartManualInput,

		/// <summary>TODO: comment</summary>
		flowChartManualOperation,

		/// <summary>TODO: comment</summary>
		flowChartConnector,

		/// <summary>TODO: comment</summary>
		flowChartOffpageConnector,

		/// <summary>TODO: comment</summary>
		flowChartPunchedCard,

		/// <summary>TODO: comment</summary>
		flowChartPunchedTape,

		/// <summary>TODO: comment</summary>
		flowChartSummingJunction,

		/// <summary>TODO: comment</summary>
		flowChartOr,

		/// <summary>TODO: comment</summary>
		flowChartCollate,

		/// <summary>TODO: comment</summary>
		flowChartSort,

		/// <summary>TODO: comment</summary>
		flowChartExtract,

		/// <summary>TODO: comment</summary>
		flowChartMerge,

		/// <summary>TODO: comment</summary>
		flowChartOnlineStorage,

		/// <summary>TODO: comment</summary>
		flowChartDelay,

		/// <summary>TODO: comment</summary>
		flowChartMagneticTape,

		/// <summary>TODO: comment</summary>
		flowChartMagneticDisk,

		/// <summary>TODO: comment</summary>
		flowChartMagneticDrum,

		/// <summary>TODO: comment</summary>
		flowChartDisplay

	};

	/// <summary>TODO: comment</summary>
	public enum StarAndBannerShapes
	{

		/// <summary>TODO: comment</summary>
		irregularSeal1,

		/// <summary>TODO: comment</summary>
		irregularSeal2,

		/// <summary>TODO: comment</summary>
		star4,

		/// <summary>TODO: comment</summary>
		star5,

		/// <summary>TODO: comment</summary>
		star6,

		/// <summary>TODO: comment</summary>
		star7,

		/// <summary>TODO: comment</summary>
		star8,

		/// <summary>TODO: comment</summary>
		star10,

		/// <summary>TODO: comment</summary>
		star12,

		/// <summary>TODO: comment</summary>
		star16,

		/// <summary>TODO: comment</summary>
		star24,

		/// <summary>TODO: comment</summary>
		star32,

		/// <summary>TODO: comment</summary>
		ribbon,

		/// <summary>TODO: comment</summary>
		ribbon2,

		/// <summary>TODO: comment</summary>
		ellipseRibbon,

		/// <summary>TODO: comment</summary>
		ellipseRibbon2,

		/// <summary>TODO: comment</summary>
		verticalScroll,

		/// <summary>TODO: comment</summary>
		horizontalScroll,

		/// <summary>TODO: comment</summary>
		wave,

		/// <summary>TODO: comment</summary>
		doubleWave

	};

	/// <summary>TODO: comment</summary>
	public enum CalloutShapes
	{

		/// <summary>TODO: comment</summary>
		wedgeRectCallout,

		/// <summary>TODO: comment</summary>
		wedgeRoundRectCallout,

		/// <summary>TODO: comment</summary>
		wedgeEllipseCallout,

		/// <summary>TODO: comment</summary>
		cloudCallout,

		/// <summary>TODO: comment</summary>
		borderCallout1,

		/// <summary>TODO: comment</summary>
		borderCallout2,

		/// <summary>TODO: comment</summary>
		borderCallout3,

		/// <summary>TODO: comment</summary>
		accentCallout1,

		/// <summary>TODO: comment</summary>
		accentCallout2,

		/// <summary>TODO: comment</summary>
		accentCallout3,

		/// <summary>TODO: comment</summary>
		callout1,

		/// <summary>TODO: comment</summary>
		callout2,

		/// <summary>TODO: comment</summary>
		callout3,

		/// <summary>TODO: comment</summary>
		accentBorderCallout1,

		/// <summary>TODO: comment</summary>
		accentBorderCallout2,

		/// <summary>TODO: comment</summary>
		accentBorderCallout3

	};

	/// <summary>Text alignment of a Paragraph</summary>
	public enum Alignment
	{

		/// <summary>Align Paragraph to the left</summary>
		left,

		/// <summary>Align Paragraph as centered</summary>
		center,

		/// <summary>Align Paragraph to the right</summary>
		right,

		/// <summary>(Justified) Align Paragraph to both the left and right margins, adding extra space between content as necessary</summary>
		both

	};

	/// <summary>TODO: comment</summary>
	public enum Direction
	{

		/// <summary>TODO: comment</summary>
		LeftToRight,

		/// <summary>TODO: comment</summary>
		RightToLeft

	};

	/// <summary>Text types in a Run</summary>
	public enum RunTextType
	{

		/// <summary>System.String</summary>
		Text,

		/// <summary>System.String</summary>
		DelText

	}

	/// <summary>TODO: comment</summary>
	public enum LineSpacingType
	{

		/// <summary>TODO: comment</summary>
		Line,

		/// <summary>TODO: comment</summary>
		Before,

		/// <summary>TODO: comment</summary>
		After

	}

	/// <summary>TODO: comment</summary>
	public enum LineSpacingTypeAuto
	{

		/// <summary>TODO: comment</summary>
		AutoBefore,

		/// <summary>TODO: comment</summary>
		AutoAfter,

		/// <summary>TODO: comment</summary>
		Auto,

		/// <summary>TODO: comment</summary>
		None

	}

	/// <summary>Cell margin for all sides of the table cell</summary>
	public enum TableCellMarginType
	{

		/// <summary>The left cell margin</summary>
		left,

		/// <summary>The right cell margin</summary>
		right,

		/// <summary>The bottom cell margin</summary>
		bottom,

		/// <summary>The top cell margin</summary>
		top

	}

	/// <summary>TODO: comment</summary>
	public enum HeadingType
	{

		/// <summary>TODO: comment</summary>
		[Description("Heading1")]
		Heading1,

		/// <summary>TODO: comment</summary>
		[Description("Heading2")]
		Heading2,

		/// <summary>TODO: comment</summary>
		[Description("Heading3")]
		Heading3,

		/// <summary>TODO: comment</summary>
		[Description("Heading4")]
		Heading4,

		/// <summary>TODO: comment</summary>
		[Description("Heading5")]
		Heading5,

		/// <summary>TODO: comment</summary>
		[Description("Heading6")]
		Heading6,

		/// <summary>TODO: comment</summary>
		[Description("Heading7")]
		Heading7,

		/// <summary>TODO: comment</summary>
		[Description("Heading8")]
		Heading8,

		/// <summary>TODO: comment</summary>
		[Description("Heading9")]
		Heading9,

	}

	/// <summary>TODO: comment</summary>
	public enum TextDirection
	{

		/// <summary>TODO: comment</summary>
		btLr,

		/// <summary>TODO: comment</summary>
		right

	};

	/// <summary>Represents the switches set on a TOC</summary>
	[Flags]
	public enum TableOfContentsSwitches
	{

		/// <summary>TODO: comment</summary>
		None = 0 << 0,

		/// <summary>TODO: comment</summary>
		[Description("\\a")]
		A = 1 << 0,

		/// <summary>TODO: comment</summary>
		[Description("\\b")]
		B = 1 << 1,

		/// <summary>TODO: comment</summary>
		[Description("\\c")]
		C = 1 << 2,

		/// <summary>TODO: comment</summary>
		[Description("\\d")]
		D = 1 << 3,

		/// <summary>TODO: comment</summary>
		[Description("\\f")]
		F = 1 << 4,

		/// <summary>TODO: comment</summary>
		[Description("\\h")]
		H = 1 << 5,

		/// <summary>TODO: comment</summary>
		[Description("\\l")]
		L = 1 << 6,

		/// <summary>TODO: comment</summary>
		[Description("\\n")]
		N = 1 << 7,

		/// <summary>TODO: comment</summary>
		[Description("\\o")]
		O = 1 << 8,

		/// <summary>TODO: comment</summary>
		[Description("\\p")]
		P = 1 << 9,

		/// <summary>TODO: comment</summary>
		[Description("\\s")]
		S = 1 << 10,

		/// <summary>TODO: comment</summary>
		[Description("\\t")]
		T = 1 << 11,

		/// <summary>TODO: comment</summary>
		[Description("\\u")]
		U = 1 << 12,

		/// <summary>TODO: comment</summary>
		[Description("\\w")]
		W = 1 << 13,

		/// <summary>TODO: comment</summary>
		[Description("\\x")]
		X = 1 << 14,

		/// <summary>TODO: comment</summary>
		[Description("\\z")]
		Z = 1 << 15

	}

	#endregion

	#endregion

	#region Interfaces

	interface IContentContainer
	{

		ReadOnlyCollection<Content> Paragraphs
		{
			get;
		}

	}

	/// <summary>TODO: comment</summary>
	public interface IParagraphContainer
	{

		/// <summary>TODO: comment</summary>
		ReadOnlyCollection<Paragraph> Paragraphs
		{
			get;
		}

	}

	#endregion

	#region Classes

	#region Internal

	internal static class Extensions
	{

		internal static string ToHex(this Color source)
		{
			byte
				red = source.R,
				green = source.G,
				blue = source.B;
			string
				redHex = red.ToString("X"),
				blueHex = blue.ToString("X"),
				greenHex = green.ToString("X");
			if (redHex.Length < 2) redHex = string.Concat("0", redHex);
			if (blueHex.Length < 2) blueHex = string.Concat("0", blueHex);
			if (greenHex.Length < 2) greenHex = string.Concat("0", greenHex);
			return string.Format("{0}{1}{2}", redHex, greenHex, blueHex);
		}

		public static void Flatten(this XElement e, XName name, List<XElement> flat)
		{
			// Add this element (without its children) to the flat list
			XElement clone = CloneElement(e);
			clone.Elements().Remove();
			// Filter elements using XName
			if (clone.Name == name) flat.Add(clone);
			// Process the children, filter elements using XName
			if (e.HasElements) foreach (XElement elem in e.Elements(name)) elem.Flatten(name, flat);
		}

		static XElement CloneElement(XElement element)
		{
			return new XElement(element.Name, element.Attributes(), element.Nodes().Select(
				n =>
				{
					XElement e = n as XElement;
					if (e != null) return CloneElement(e);
					return n;
				}
			));
		}

		public static string GetAttribute(this XElement el, XName name, string defaultValue = "")
		{
			var attr = el.Attribute(name);
			if (attr != null) return attr.Value;
			return defaultValue;
		}

		/// <summary>Sets margin for all the pages in a Dox document in Inches. Written by Shashwat Tripathi.</summary>
		/// <param name="document"></param>
		/// <param name="top">Margin from the Top. Leave -1 for no change</param>
		/// <param name="bottom">Margin from the Bottom. Leave -1 for no change</param>
		/// <param name="right">Margin from the Right. Leave -1 for no change</param>
		/// <param name="left">Margin from the Left. Leave -1 for no change</param>
		public static void SetMargin(this DocX document, float top, float bottom, float right, float left)
		{
			XNamespace ab = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
			var tempElement = document.PageLayout.Xml.Descendants(ab + "pgMar");
			var e = tempElement.GetEnumerator();
			foreach (var item in tempElement)
			{
				if (left != -1) item.SetAttributeValue(ab + "left", (1440 * left) / 1);
				if (right != -1) item.SetAttributeValue(ab + "right", (1440 * right) / 1);
				if (top != -1) item.SetAttributeValue(ab + "top", (1440 * top) / 1);
				if (bottom != -1) item.SetAttributeValue(ab + "bottom", (1440 * bottom) / 1);
			}
		}

	}

	internal static class HelperFunctions
	{

		public const string DOCUMENT_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

		public const string TEMPLATE_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

		public static bool IsNullOrWhiteSpace(this string value)
		{
			if (value == null) return true;
			return string.IsNullOrEmpty(value.Trim());
		}

		/// <summary>Checks whether 'toCheck' has all children that 'desired' has and values of 'val' attributes are the same</summary>
		/// <param name="desired">TODO: comment</param>
		/// <param name="toCheck">TODO: comment</param>
		/// <param name="fo">Matching options whether check if desired attributes are inder a, or a has exactly and only these attributes as b has.</param>
		/// <returns></returns>
		internal static bool ContainsEveryChildOf(XElement desired, XElement toCheck, MatchFormattingOptions fo)
		{
			foreach (XElement e in desired.Elements())
			{
				// If a formatting property has the same name and 'val' attribute's value, its considered to be equivalent
				if (!toCheck.Elements(e.Name).Where(bElement => bElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName)) == e.GetAttribute(XName.Get("val", DocX.w.NamespaceName))).Any()) return false;
			}
			// If the formatting has to be exact, no additionaly formatting must exist
			if (fo == MatchFormattingOptions.ExactMatch) return desired.Elements().Count() == toCheck.Elements().Count();
			return true;
		}

		internal static void CreateRelsPackagePart(DocX Document, Uri uri)
		{
			PackagePart pp = Document.package.CreatePart(uri, "application/vnd.openxmlformats-package.relationships+xml", CompressionOption.Maximum);
			using (TextWriter tw = new StreamWriter(new PackagePartStream(pp.GetStream())))
			{
				XDocument d = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("Relationships", DocX.rel.NamespaceName)));
				var root = d.Root;
				d.Save(tw);
			}
		}

		internal static int GetSize(XElement Xml)
		{
			switch (Xml.Name.LocalName)
			{
				case "tab":
				case "br":
					return 1;
				case "t":
					goto case "delText";
				case "delText":
					return Xml.Value.Length;
				case "tr":
				case "tc":
					goto case "br";
				default:
					return 0;
			}
		}

		internal static string GetText(XElement e)
		{
			StringBuilder sb = new StringBuilder();
			GetTextRecursive(e, ref sb);
			return sb.ToString();
		}

		internal static void GetTextRecursive(XElement Xml, ref StringBuilder sb)
		{
			sb.Append(ToText(Xml));
			if (Xml.HasElements) foreach (XElement e in Xml.Elements()) GetTextRecursive(e, ref sb);
		}

		internal static List<FormattedText> GetFormattedText(XElement e)
		{
			List<FormattedText> alist = new List<FormattedText>();
			GetFormattedTextRecursive(e, ref alist);
			return alist;
		}

		internal static void GetFormattedTextRecursive(XElement Xml, ref List<FormattedText> alist)
		{
			FormattedText
				ft = ToFormattedText(Xml),
				last = null;
			if (ft != null)
			{
				if (alist.Count() > 0) last = alist.Last();
				if (last != null && last.CompareTo(ft) == 0)
				{
					// Update text of last entry
					last.text += ft.text;
				}
				else
				{
					if (last != null) ft.index = last.index + last.text.Length;
					alist.Add(ft);
				}
			}
			if (Xml.HasElements) foreach (XElement e in Xml.Elements()) GetFormattedTextRecursive(e, ref alist);
		}

		internal static FormattedText ToFormattedText(XElement e)
		{
			// The text representation of e
			string text = ToText(e);
			if (text == string.Empty) return null;
			// e is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it
			while (!e.Name.Equals(XName.Get("r", DocX.w.NamespaceName)) && !e.Name.Equals(XName.Get("tabs", DocX.w.NamespaceName))) e = e.Parent;
			// e is a w:r element, lets find the rPr element
			XElement rPr = e.Element(XName.Get("rPr", DocX.w.NamespaceName));
			FormattedText ft = new FormattedText();
			ft.text = text;
			ft.index = 0;
			ft.formatting = null;
			// Return text with formatting
			if (rPr != null) ft.formatting = Formatting.Parse(rPr);
			return ft;
		}

		internal static string ToText(XElement e)
		{
			switch (e.Name.LocalName)
			{
				case "tab":
					return "\t";
				case "br":
					return "\n";
				case "t":
					goto case "delText";
				case "delText":
					if (e.Parent != null && e.Parent.Name.LocalName == "r")
					{
						XElement run = e.Parent;
						var rPr = run.Elements().FirstOrDefault(a => a.Name.LocalName == "rPr");
						if (rPr != null)
						{
							var caps = rPr.Elements().FirstOrDefault(a => a.Name.LocalName == "caps");
							if (caps != null) return e.Value.ToUpper();
						}
					}
					return e.Value;
				case "tr":
					goto case "br";
				case "tc":
					goto case "tab";
				default:
					return string.Empty;
			}
		}

		internal static XElement CloneElement(XElement element)
		{
			return new XElement(element.Name, element.Attributes(), element.Nodes().Select(
				n =>
				{
					XElement e = n as XElement;
					if (e != null) return CloneElement(e);
					return n;
				}
			));
		}

		internal static PackagePart CreateOrGetSettingsPart(Package package)
		{
			PackagePart settingsPart;
			Uri settingsUri = new Uri("/word/settings.xml", UriKind.Relative);
			if (!package.PartExists(settingsUri))
			{
				settingsPart = package.CreatePart(settingsUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", CompressionOption.Maximum);
				PackagePart mainDocumentPart = package.GetParts().Single(p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase));
				mainDocumentPart.CreateRelationship(settingsUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
				XDocument settings = XDocument.Parse(@"<?xml version='1.0' encoding='utf-8' standalone='yes'?><w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'><w:zoom w:percent='100' /><w:defaultTabStop w:val='720' /><w:characterSpacingControl w:val='doNotCompress' /><w:compat /><w:rsids><w:rsidRoot w:val='00217F62' /><w:rsid w:val='001915A3' /><w:rsid w:val='00217F62' /><w:rsid w:val='00A906D8' /><w:rsid w:val='00AB5A74' /><w:rsid w:val='00F071AE' /></w:rsids><m:mathPr><m:mathFont m:val='Cambria Math' /><m:brkBin m:val='before' /><m:brkBinSub m:val='--' /><m:smallFrac m:val='off' /><m:dispDef /><m:lMargin m:val='0' /><m:rMargin m:val='0' /><m:defJc m:val='centerGroup' /><m:wrapIndent m:val='1440' /><m:intLim m:val='subSup' /><m:naryLim m:val='undOvr' /></m:mathPr><w:themeFontLang w:val='en-IE' w:bidi='ar-SA' /><w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' /><w:shapeDefaults><o:shapedefaults v:ext='edit' spidmax='2050' /><o:shapelayout v:ext='edit'><o:idmap v:ext='edit' data='1' /></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val='.' /><w:listSeparator w:val=',' /></w:settings>");
				XElement themeFontLang = settings.Root.Element(XName.Get("themeFontLang", DocX.w.NamespaceName));
				themeFontLang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);
				// Save the settings document
				using (TextWriter tw = new StreamWriter(new PackagePartStream(settingsPart.GetStream()))) settings.Save(tw);
			}
			else settingsPart = package.GetPart(settingsUri);
			return settingsPart;
		}

		internal static void CreateCustomPropertiesPart(DocX document)
		{
			PackagePart customPropertiesPart = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum);
			XDocument customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("Properties", DocX.customPropertiesSchema.NamespaceName), new XAttribute(XNamespace.Xmlns + "vt", DocX.customVTypesSchema)));
			using (TextWriter tw = new StreamWriter(new PackagePartStream(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))) customPropDoc.Save(tw, SaveOptions.None);
			document.package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
		}

		internal static XDocument DecompressXMLResource(string manifest_resource_name)
		{
			// XDocument to load the compressed Xml resource into
			XDocument document;
			// Get a reference to the executing assembly
			Assembly assembly = Assembly.GetExecutingAssembly();
			// Open a Stream to the embedded resource
			Stream stream = assembly.GetManifestResourceStream(manifest_resource_name);
			// Decompress the embedded resource
			// Load this decompressed embedded resource into an XDocument using a TextReader
			using (GZipStream zip = new GZipStream(stream, CompressionMode.Decompress))
			using (TextReader sr = new StreamReader(zip))
			{
				document = XDocument.Load(sr);
			}
			// Return the decompressed Xml as an XDocument
			return document;
		}

		/// <summary>If this document does not contain a /word/numbering.xml add the default one generated by Microsoft Word when the default bullet, numbered and multilevel lists are added to a blank document</summary>
		/// <param name="package">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		internal static XDocument AddDefaultNumberingXml(Package package)
		{
			XDocument numberingDoc;
			// Create the main document part for this package
			PackagePart wordNumbering = package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);
			numberingDoc = DecompressXMLResource("Novacode.Resources.numbering.xml.gz");
			// Save /word/numbering.xml
			using (TextWriter tw = new StreamWriter(new PackagePartStream(wordNumbering.GetStream(FileMode.Create, FileAccess.Write)))) numberingDoc.Save(tw, SaveOptions.None);
			PackagePart mainDocumentPart = package.GetParts().Single(p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase));
			mainDocumentPart.CreateRelationship(wordNumbering.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
			return numberingDoc;
		}

		/// <summary>If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word</summary>
		/// <param name="package">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		internal static XDocument AddDefaultStylesXml(Package package)
		{
			XDocument stylesDoc;
			// Create the main document part for this package
			PackagePart word_styles = package.CreatePart(new Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
			stylesDoc = DecompressXMLResource("Novacode.Resources.default_styles.xml.gz");
			XElement lang = stylesDoc.Root.Element(XName.Get("docDefaults", DocX.w.NamespaceName)).Element(XName.Get("rPrDefault", DocX.w.NamespaceName)).Element(XName.Get("rPr", DocX.w.NamespaceName)).Element(XName.Get("lang", DocX.w.NamespaceName));
			lang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);
			// Save /word/styles.xml
			using (TextWriter tw = new StreamWriter(new PackagePartStream(word_styles.GetStream(FileMode.Create, FileAccess.Write)))) stylesDoc.Save(tw, SaveOptions.None);
			PackagePart mainDocumentPart = package.GetParts().Where(p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)).Single();
			mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
			return stylesDoc;
		}

		internal static XElement CreateEdit(EditType t, DateTime edit_time, object content)
		{
			if (t == EditType.del)
			{
				foreach (object o in (IEnumerable<XElement>)content)
				{
					if (o is XElement)
					{
						XElement e = (o as XElement);
						IEnumerable<XElement> ts = e.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName));
						for (int i = 0; i < ts.Count(); i++)
						{
							XElement text = ts.ElementAt(i);
							text.ReplaceWith(new XElement(DocX.w + "delText", text.Attributes(), text.Value));
						}
					}
				}
			}
			return (new XElement(DocX.w + t.ToString(), new XAttribute(DocX.w + "id", 0), new XAttribute(DocX.w + "author", WindowsIdentity.GetCurrent().Name), new XAttribute(DocX.w + "date", edit_time), content));
		}

		internal static XElement CreateTable(int rowCount, int columnCount)
		{
			int[] columnWidths = new int[columnCount];
			for (int i = 0; i < columnCount; i++) columnWidths[i] = 2310;
			return CreateTable(rowCount, columnWidths);
		}

		internal static XElement CreateTable(int rowCount, int[] columnWidths)
		{
			XElement newTable = new XElement(XName.Get("tbl", DocX.w.NamespaceName), new XElement(XName.Get("tblPr", DocX.w.NamespaceName), new XElement(XName.Get("tblStyle", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "TableGrid")), new XElement(XName.Get("tblW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "5000"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "auto")), new XElement(XName.Get("tblLook", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "04A0"))));
			for (int i = 0; i < rowCount; i++)
			{
				XElement row = new XElement(XName.Get("tr", DocX.w.NamespaceName));
				for (int j = 0; j < columnWidths.Length; j++)
				{
					XElement cell = CreateTableCell();
					row.Add(cell);
				}
				newTable.Add(row);
			}
			return newTable;
		}

		/// <summary>Create and return a cell of a table</summary>
		internal static XElement CreateTableCell(double w = 2310)
		{
			return new XElement(XName.Get("tc", DocX.w.NamespaceName), new XElement(XName.Get("tcPr", DocX.w.NamespaceName), new XElement(XName.Get("tcW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), w), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))), new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName))));
		}

		internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
		{
			if (list.NumId == 0) list.CreateNewNumberingNumId(level, listType, startNumber, continueNumbering);
			// I see no reason why you shouldn't be able to insert an empty element. It simplifies tasks such as populating an item from html.
			if (listText != null)
			{
				var newParagraphSection = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("numPr", DocX.w.NamespaceName), new XElement(XName.Get("ilvl", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", level)), new XElement(XName.Get("numId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", list.NumId)))), new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("t", DocX.w.NamespaceName), listText)));
				if (trackChanges) newParagraphSection = CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);
				if (startNumber == null)
				{
					list.AddItem(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph));
				}
				else
				{
					list.AddItemWithStartValue(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph), (int)startNumber);
				}
			}
			return list;
		}

		internal static void RenumberIDs(DocX document)
		{
			IEnumerable<XAttribute> trackerIDs = (from d in document.mainDoc.Descendants() where d.Name.LocalName == "ins" || d.Name.LocalName == "del" select d.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")));
			for (int i = 0; i < trackerIDs.Count(); i++) trackerIDs.ElementAt(i).Value = i.ToString();
		}

		internal static Paragraph GetFirstParagraphEffectedByInsert(DocX document, int index)
		{
			// This document contains no Paragraphs and insertion is at index 0
			if (document.paragraphLookup.Keys.Count() == 0 && index == 0) return null;
			foreach (int paragraphEndIndex in document.paragraphLookup.Keys) if (paragraphEndIndex >= index) return document.paragraphLookup[paragraphEndIndex];
			throw new ArgumentOutOfRangeException();
		}

		internal static List<XElement> FormatInput(string text, XElement rPr)
		{
			List<XElement> newRuns = new List<XElement>();
			XElement
				tabRun = new XElement(DocX.w + "tab"),
				breakRun = new XElement(DocX.w + "br");
			StringBuilder sb = new StringBuilder();
			// I dont wanna get an exception if text == null, so just return empty list
			if (string.IsNullOrEmpty(text)) return newRuns;
			char lastChar = '\0';
			foreach (char c in text)
			{
				switch (c)
				{
					case '\t':
						if (sb.Length > 0)
						{
							XElement t = new XElement(DocX.w + "t", sb.ToString());
							Text.PreserveSpace(t);
							newRuns.Add(new XElement(DocX.w + "r", rPr, t));
							sb = new StringBuilder();
						}
						newRuns.Add(new XElement(DocX.w + "r", rPr, tabRun));
						break;
					case '\r':
						if (sb.Length > 0)
						{
							XElement t = new XElement(DocX.w + "t", sb.ToString());
							Text.PreserveSpace(t);
							newRuns.Add(new XElement(DocX.w + "r", rPr, t));
							sb = new StringBuilder();
						}
						newRuns.Add(new XElement(DocX.w + "r", rPr, breakRun));
						break;
					case '\n':
						if (lastChar == '\r') break;
						if (sb.Length > 0)
						{
							XElement t = new XElement(DocX.w + "t", sb.ToString());
							Text.PreserveSpace(t);
							newRuns.Add(new XElement(DocX.w + "r", rPr, t));
							sb = new StringBuilder();
						}
						newRuns.Add(new XElement(DocX.w + "r", rPr, breakRun));
						break;
					default:
						sb.Append(c);
						break;
				}
				lastChar = c;
			}
			if (sb.Length > 0)
			{
				XElement t = new XElement(DocX.w + "t", sb.ToString());
				Text.PreserveSpace(t);
				newRuns.Add(new XElement(DocX.w + "r", rPr, t));
			}
			return newRuns;
		}

		internal static XElement[] SplitParagraph(Paragraph p, int index)
		{
			// In this case edit dosent really matter, you have a choice
			Run r = p.GetFirstRunEffectedByEdit(index, EditType.ins);
			XElement[] split;
			XElement before, after;
			if (r.Xml.Parent.Name.LocalName == "ins")
			{
				split = p.SplitEdit(r.Xml.Parent, index, EditType.ins);
				before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
				after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
			}
			else if (r.Xml.Parent.Name.LocalName == "del")
			{
				split = p.SplitEdit(r.Xml.Parent, index, EditType.del);
				before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
				after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
			}
			else
			{
				split = Run.SplitRun(r, index);
				before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[0]);
				after = new XElement(p.Xml.Name, p.Xml.Attributes(), split[1], r.Xml.ElementsAfterSelf());
			}
			if (before.Elements().Count() == 0) before = null;
			if (after.Elements().Count() == 0) after = null;
			return new XElement[] { before, after };
		}

		/// <summary>Bug found and fixed by trnilse. To see the change, please compare this release to the previous release using TFS compare</summary>
		/// <param name="streamOne"></param>
		/// <param name="streamTwo"></param>
		/// <returns></returns>
		internal static bool IsSameFile(Stream streamOne, Stream streamTwo)
		{
			int file1byte,
				file2byte;
			// Return false to indicate files are different
			if (streamOne.Length != streamTwo.Length) return false;
			// Read and compare a byte from each file until either a non-matching set of bytes is found or until the end of file1 is reached
			do
			{
				// Read one byte from each file
				file1byte = streamOne.ReadByte();
				file2byte = streamTwo.ReadByte();
			}
			while ((file1byte == file2byte) && (file1byte != -1));
			// Return the success of the comparison. "file1byte" is equal to "file2byte" at this point only if the files are the same.
			streamOne.Position = 0;
			streamTwo.Position = 0;
			return ((file1byte - file2byte) == 0);
		}

		internal static UnderlineStyle GetUnderlineStyle(string underlineStyle)
		{
			switch (underlineStyle)
			{
				case "single":
					return UnderlineStyle.singleLine;
				case "double":
					return UnderlineStyle.doubleLine;
				case "thick":
					return UnderlineStyle.thick;
				case "dotted":
					return UnderlineStyle.dotted;
				case "dottedHeavy":
					return UnderlineStyle.dottedHeavy;
				case "dash":
					return UnderlineStyle.dash;
				case "dashedHeavy":
					return UnderlineStyle.dashedHeavy;
				case "dashLong":
					return UnderlineStyle.dashLong;
				case "dashLongHeavy":
					return UnderlineStyle.dashLongHeavy;
				case "dotDash":
					return UnderlineStyle.dotDash;
				case "dashDotHeavy":
					return UnderlineStyle.dashDotHeavy;
				case "dotDotDash":
					return UnderlineStyle.dotDotDash;
				case "dashDotDotHeavy":
					return UnderlineStyle.dashDotDotHeavy;
				case "wave":
					return UnderlineStyle.wave;
				case "wavyHeavy":
					return UnderlineStyle.wavyHeavy;
				case "wavyDouble":
					return UnderlineStyle.wavyDouble;
				case "words":
					return UnderlineStyle.words;
				default:
					return UnderlineStyle.none;
			}
		}

	}

	internal class Text : DocXElement
	{

		private int startIndex;

		private int endIndex;

		private string text;

		/// <summary>Gets the start index of this Text (text length before this text)</summary>
		public int StartIndex
		{
			get
			{
				return startIndex;
			}
		}

		/// <summary>Gets the end index of this Text (text length before this text + this texts length)</summary>
		public int EndIndex
		{
			get
			{
				return endIndex;
			}
		}

		/// <summary>The text value of this text element</summary>
		public string Value
		{
			get
			{
				return text;
			}
		}

		internal Text(DocX document, XElement xml, int startIndex) : base(document, xml)
		{
			this.startIndex = startIndex;
			switch (Xml.Name.LocalName)
			{
				case "t":
					goto case "delText";
				case "delText":
					endIndex = startIndex + xml.Value.Length;
					text = xml.Value;
					break;
				case "br":
					text = "\n";
					endIndex = startIndex + 1;
					break;
				case "tab":
					text = "\t";
					endIndex = startIndex + 1;
					break;
			}
		}

		internal static XElement[] SplitText(Text t, int index)
		{
			if (index < t.startIndex || index > t.EndIndex) throw new ArgumentOutOfRangeException(nameof(index));
			XElement splitLeft = null, splitRight = null;
			if (t.Xml.Name.LocalName == "t" || t.Xml.Name.LocalName == "delText")
			{
				// The origional text element, now containing only the text before the index point
				splitLeft = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(0, index - t.startIndex));
				if (splitLeft.Value.Length == 0) splitLeft = null;
				else PreserveSpace(splitLeft);
				// The origional text element, now containing only the text after the index point
				splitRight = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(index - t.startIndex, t.Xml.Value.Length - (index - t.startIndex)));
				if (splitRight.Value.Length == 0) splitRight = null;
				else PreserveSpace(splitRight);
			}
			else
			{
				if (index == t.EndIndex) splitLeft = t.Xml;
				else splitRight = t.Xml;
			}
			return (new XElement[] { splitLeft, splitRight });
		}

		/// <summary>If a text element or delText element, starts or ends with a space, it must have the attribute space, otherwise it must not have it</summary>
		/// <param name="e">The (t or delText) element check</param>
		public static void PreserveSpace(XElement e)
		{
			// PreserveSpace should only be used on (t or delText) elements
			if (!e.Name.Equals(DocX.w + "t") && !e.Name.Equals(DocX.w + "delText")) throw new ArgumentException("SplitText can only split elements of type t or delText", "e");
			// Check if this w:t contains a space atribute
			XAttribute space = e.Attributes().Where(a => a.Name.Equals(XNamespace.Xml + "space")).SingleOrDefault();
			// This w:t's text begins or ends with whitespace
			if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
			{
				// If this w:t contains no space attribute, add one
				if (space == null) e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
			}
			// This w:t's text does not begin or end with a space
			else
			{
				// If this w:r contains a space attribute, remove it
				if (space != null) space.Remove();
			}
		}

	}

	internal static class XElementHelpers
	{

		/// <summary>Get value from XElement and convert it to enum</summary>
		/// <typeparam name="T">Enum type</typeparam>
		internal static T GetValueToEnum<T>(XElement element)
		{
			if (element == null) throw new ArgumentNullException(nameof(element));
			string value = element.Attribute(XName.Get("val")).Value;
			foreach (T e in Enum.GetValues(typeof(T)))
			{
				FieldInfo fi = typeof(T).GetField(e.ToString());
				if (fi.GetCustomAttributes(typeof(XmlNameAttribute), false).Count() == 0) throw new Exception(string.Format("Attribute 'XmlNameAttribute' is not assigned to {0} fields!", typeof(T).Name));
				XmlNameAttribute a = (XmlNameAttribute)fi.GetCustomAttributes(typeof(XmlNameAttribute), false).First();
				if (a.XmlName == value) return e;
			}
			throw new ArgumentException("Invalid element value!");
		}

		/// <summary>Convert value to xml string and set it into XElement</summary>
		/// <typeparam name="T">Enum type</typeparam> 
		internal static void SetValueFromEnum<T>(XElement element, T value)
		{
			if (element == null) throw new ArgumentNullException(nameof(element));
			element.Attribute(XName.Get("val")).Value = GetXmlNameFromEnum<T>(value);
		}

		/// <summary>Return xml string for this value</summary>
		/// <typeparam name="T">Enum type</typeparam>
		internal static string GetXmlNameFromEnum<T>(T value)
		{
			if (value == null) throw new ArgumentNullException(nameof(value));
			FieldInfo fi = typeof(T).GetField(value.ToString());
			if (fi.GetCustomAttributes(typeof(XmlNameAttribute), false).Count() == 0) throw new Exception(string.Format("Attribute 'XmlNameAttribute' is not assigned to {0} fields!", typeof(T).Name));
			XmlNameAttribute a = (XmlNameAttribute)fi.GetCustomAttributes(typeof(XmlNameAttribute), false).First();
			return a.XmlName;
		}

	}

	/// <summary>This attribute applied to enum's fields for definition their's real xml names in DocX file</summary>
	/// <example>
	///		public enum MyEnum
	///		{
	///			[XmlName("one")] // This means, that xml element has 'val="one"'
	///			ValueOne,
	///			[XmlName("two")] // This means, that xml element has 'val="two"'
	///			ValueTwo
	///		}
	/// </example>
	[AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
	internal sealed class XmlNameAttribute : Attribute
	{

		/// <summary>Real xml name</summary>
		public string XmlName
		{
			get;
			private set;
		}

		public XmlNameAttribute(string xmlName)
		{
			XmlName = xmlName;
		}

	}

	#endregion

	#region Public

	/// <summary>Axis base class</summary>
	public abstract class Axis
	{

		/// <summary>ID of this Axis</summary>
		public string Id
		{
			get
			{
				return Xml.Element(XName.Get("axId", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value;
			}
		}

		/// <summary>Return true if this axis is visible</summary>
		public bool IsVisible
		{
			get
			{
				return "0".Equals(Xml.Element(XName.Get("delete", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value);
			}
			set
			{
				Xml.Element(XName.Get("delete", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = value ? "0" : "1";
			}
		}

		/// <summary>Axis xml element</summary>
		internal XElement Xml
		{
			get;
			set;
		}

		internal Axis(XElement xml)
		{
			Xml = xml;
		}

		/// <summary>Конструктор</summary>
		/// <param name="id">TODO: comment</param>
		public Axis(string id)
		{
		}

	}

	/// <summary>Represents Category Axes</summary>
	public class CategoryAxis : Axis
	{

		internal CategoryAxis(XElement xml) : base(xml)
		{
		}

		/// <summary>TODO: comment</summary>
		/// <param name="id">TODO: comment</param>
		public CategoryAxis(string id) : base(id)
		{
			Xml = XElement.Parse(string.Format(@"<c:catAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:axId val=""{0}""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""b""/><c:majorTickMark val=""out""/><c:minorTickMark val=""none""/><c:tickLblPos val=""nextTo""/><c:crossAx val=""154227840""/><c:crosses val=""autoZero""/><c:auto val=""1""/><c:lblAlgn val=""ctr""/><c:lblOffset val=""100""/><c:noMultiLvlLbl val=""0""/></c:catAx>", id));
		}

	}

	/// <summary>Represents Values Axes</summary>
	public class ValueAxis : Axis
	{

		internal ValueAxis(XElement xml) : base(xml)
		{
		}

		/// <summary>TODO: comment</summary>
		/// <param name="id">TODO: comment</param>
		public ValueAxis(string id) : base(id)
		{
			Xml = XElement.Parse(string.Format(@"<c:valAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:axId val=""{0}""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""l""/><c:numFmt sourceLinked=""0"" formatCode=""General""/><c:majorGridlines/><c:majorTickMark val=""out""/><c:minorTickMark val=""none""/><c:tickLblPos val=""nextTo""/><c:crossAx val=""148921728""/><c:crosses val=""autoZero""/><c:crossBetween val=""between""/></c:valAx>", id));
		}

	}

	/// <summary>This element contains the 2-D bar or column series on this chart. 21.2.2.16 barChart (Bar Charts).</summary>
	public class BarChart : Chart
	{

		/// <summary>Specifies the possible directions for a bar chart</summary>
		public BarDirection BarDirection
		{
			get
			{
				return XElementHelpers.GetValueToEnum<BarDirection>(ChartXml.Element(XName.Get("barDir", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(ChartXml.Element(XName.Get("barDir", DocX.c.NamespaceName)), value);
			}
		}

		/// <summary>Specifies the possible groupings for a bar chart</summary>
		public BarGrouping BarGrouping
		{
			get
			{
				return XElementHelpers.GetValueToEnum<BarGrouping>(ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)), value);
			}
		}

		/// <summary>Specifies that its contents contain a percentage between 0% and 500%</summary>
		public int GapWidth
		{
			get
			{
				return Convert.ToInt32(ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value);
			}
			set
			{
				if ((value < 1) || (value > 500)) throw new ArgumentException("GapWidth lay between 0% and 500%!");
				ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = value.ToString();
			}
		}

		/// <summary>TODO: comment</summary>
		/// <returns>TODO: comment</returns>
		protected override XElement CreateChartXml()
		{
			return XElement.Parse(@"<c:barChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:barDir val=""col""/><c:grouping val=""clustered""/><c:gapWidth val=""150""/></c:barChart>");
		}

	}

	/// <summary>Represents every Chart in this document</summary>
	public abstract class Chart
	{

		/// <summary>TODO: comment</summary>
		protected XElement ChartXml
		{
			get;
			private set;
		}

		/// <summary>TODO: comment</summary>
		protected XElement ChartRootXml
		{
			get;
			private set;
		}

		/// <summary>The xml representation of this chart</summary>
		public XDocument Xml
		{
			get;
			private set;
		}

		/// <summary>Chart's series</summary>
		public List<Series> Series
		{
			get
			{
				List<Series> series = new List<Series>();
				int indx = 1;
				foreach (XElement element in ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)))
				{
					element.Add(new XElement(XName.Get("idx", DocX.c.NamespaceName)), (indx++).ToString());
					series.Add(new Series(element));
				}
				return series;
			}
		}

		/// <summary>Return maximum count of series</summary>
		public virtual short MaxSeriesCount
		{
			get
			{
				return short.MaxValue;
			}
		}

		/// <summary>Add a new series to this chart</summary>
		public void AddSeries(Series series)
		{
			int serCount = ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)).Count();
			if (serCount == MaxSeriesCount) throw new InvalidOperationException("Maximum series for this chart is" + MaxSeriesCount.ToString() + "and have exceeded!");
			// Sourceman 16.04.2015 - Every Series needs to have an order and index element for being processed by Word 2013
			series.Xml.AddFirst(new XElement(XName.Get("order", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), (serCount + 1).ToString())));
			series.Xml.AddFirst(new XElement(XName.Get("idx", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), (serCount + 1).ToString())));
			ChartXml.Add(series.Xml);
		}

		/// <summary>Chart's legend. If legend doesn't exist property is null.</summary>
		public ChartLegend Legend
		{
			get;
			private set;
		}

		/// <summary>Add standart legend to the chart</summary>
		public void AddLegend()
		{
			AddLegend(ChartLegendPosition.Right, false);
		}

		/// <summary>Add a legend with parameters to the chart</summary>
		public void AddLegend(ChartLegendPosition position, bool overlay)
		{
			if (Legend != null) RemoveLegend();
			Legend = new ChartLegend(position, overlay);
			// Sourceman: seems to be necessary to keep track of the order of elements as defined in the schema (Word 2013)
			ChartRootXml.Element(XName.Get("plotArea", DocX.c.NamespaceName)).AddAfterSelf(Legend.Xml);
		}

		/// <summary>Remove the legend from the chart</summary>
		public void RemoveLegend()
		{
			Legend.Xml.Remove();
			Legend = null;
		}

		/// <summary>Represents the category axis</summary>
		public CategoryAxis CategoryAxis
		{
			get;
			private set;
		}

		/// <summary>Represents the values axis</summary>
		public ValueAxis ValueAxis
		{
			get;
			private set;
		}

		/// <summary>Represents existing the axis</summary>
		public virtual bool IsAxisExist
		{
			get
			{
				return true;
			}
		}

		/// <summary>Get or set 3D view for this chart</summary>
		public bool View3D
		{
			get
			{
				return ChartXml.Name.LocalName.Contains("3D");
			}
			set
			{
				if (value)
				{
					if (!View3D)
					{
						string currentName = ChartXml.Name.LocalName;
						ChartXml.Name = XName.Get(currentName.Replace("Chart", "3DChart"), DocX.c.NamespaceName);
					}
				}
				else
				{
					if (View3D)
					{
						string currentName = ChartXml.Name.LocalName;
						ChartXml.Name = XName.Get(currentName.Replace("3DChart", "Chart"), DocX.c.NamespaceName);
					}
				}
			}
		}

		/// <summary>Specifies how blank cells shall be plotted on a chart</summary>
		public DisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				return XElementHelpers.GetValueToEnum<DisplayBlanksAs>(ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)), value);
			}
		}

		/// <summary>Create an Chart for this document</summary>
		public Chart()
		{
			// Create global xml
			Xml = XDocument.Parse(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><c:roundedCorners val=""0""/><c:chart><c:autoTitleDeleted val=""0""/><c:plotVisOnly val=""1""/><c:dispBlanksAs val=""gap""/><c:showDLblsOverMax val=""0""/></c:chart></c:chartSpace>");
			// Create a real chart xml in an inheritor
			ChartXml = CreateChartXml();
			// Create result plotarea element
			XElement plotAreaXml = new XElement(XName.Get("plotArea", DocX.c.NamespaceName), new XElement(XName.Get("layout", DocX.c.NamespaceName)), ChartXml);
			// Set labels 
			XElement dLblsXml = XElement.Parse(@"<c:dLbls xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:showLegendKey val=""0""/><c:showVal val=""0""/><c:showCatName val=""0""/><c:showSerName val=""0""/><c:showPercent val=""0""/><c:showBubbleSize val=""0""/><c:showLeaderLines val=""1""/></c:dLbls>");
			ChartXml.Add(dLblsXml);
			// if axes exists, create their
			if (IsAxisExist)
			{
				CategoryAxis = new CategoryAxis("148921728");
				ValueAxis = new ValueAxis("154227840");
				XElement axIDcatXml = XElement.Parse(string.Format(@"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", CategoryAxis.Id));
				XElement axIDvalXml = XElement.Parse(string.Format(@"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", ValueAxis.Id));
				// Sourceman: seems to be necessary to keep track of the order of elements as defined in the schema (Word 2013)
				var insertPoint = ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName));
				if (insertPoint != null)
				{
					insertPoint.AddAfterSelf(axIDvalXml);
					insertPoint.AddAfterSelf(axIDcatXml);
				}
				else
				{
					ChartXml.Add(axIDcatXml);
					ChartXml.Add(axIDvalXml);
				}
				plotAreaXml.Add(CategoryAxis.Xml);
				plotAreaXml.Add(ValueAxis.Xml);
			}
			ChartRootXml = Xml.Root.Element(XName.Get("chart", DocX.c.NamespaceName));
			ChartRootXml.Element(XName.Get("autoTitleDeleted", DocX.c.NamespaceName)).AddAfterSelf(plotAreaXml);
		}

		/// <summary>An abstract method which creates the current chart xml</summary>
		protected abstract XElement CreateChartXml();

	}

	/// <summary>Represents a chart series</summary>
	public class Series
	{

		private XElement strCache;

		private XElement numCache;

		/// <summary>Series xml element</summary>
		internal XElement Xml
		{
			get;
			private set;
		}

		/// <summary>TODO: comment</summary>
		public Color Color
		{
			get
			{
				XElement colorElement = Xml.Element(XName.Get("spPr", DocX.c.NamespaceName));
				if (colorElement == null) return Color.Transparent;
				else return Color.FromArgb(int.Parse(colorElement.Element(XName.Get("solidFill", DocX.a.NamespaceName)).Element(XName.Get("srgbClr", DocX.a.NamespaceName)).Attribute(XName.Get("val")).Value, NumberStyles.HexNumber));
			}
			set
			{
				XElement colorElement = Xml.Element(XName.Get("spPr", DocX.c.NamespaceName));
				if (colorElement != null) colorElement.Remove();
				colorElement = new XElement(XName.Get("spPr", DocX.c.NamespaceName), new XElement(XName.Get("solidFill", DocX.a.NamespaceName), new XElement(XName.Get("srgbClr", DocX.a.NamespaceName), new XAttribute(XName.Get("val"), value.ToHex()))));
				Xml.Element(XName.Get("tx", DocX.c.NamespaceName)).AddAfterSelf(colorElement);
			}
		}

		internal Series(XElement xml)
		{
			Xml = xml;
			strCache = xml.Element(XName.Get("cat", DocX.c.NamespaceName)).Element(XName.Get("strRef", DocX.c.NamespaceName)).Element(XName.Get("strCache", DocX.c.NamespaceName));
			numCache = xml.Element(XName.Get("val", DocX.c.NamespaceName)).Element(XName.Get("numRef", DocX.c.NamespaceName)).Element(XName.Get("numCache", DocX.c.NamespaceName));
		}

		/// <summary>TODO: comment</summary>
		/// <param name="name">TODO: comment</param>
		public Series(string name)
		{
			strCache = new XElement(XName.Get("strCache", DocX.c.NamespaceName));
			numCache = new XElement(XName.Get("numCache", DocX.c.NamespaceName));
			Xml = new XElement(XName.Get("ser", DocX.c.NamespaceName), new XElement(XName.Get("tx", DocX.c.NamespaceName), new XElement(XName.Get("strRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), new XElement(XName.Get("strCache", DocX.c.NamespaceName), new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), "0"), new XElement(XName.Get("v", DocX.c.NamespaceName), name))))), new XElement(XName.Get("invertIfNegative", DocX.c.NamespaceName), "0"), new XElement(XName.Get("cat", DocX.c.NamespaceName), new XElement(XName.Get("strRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), strCache)), new XElement(XName.Get("val", DocX.c.NamespaceName), new XElement(XName.Get("numRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), numCache)));
		}

		/// <summary>TODO: comment</summary>
		/// <param name="list">TODO: comment</param>
		/// <param name="categoryPropertyName">TODO: comment</param>
		/// <param name="valuePropertyName">TODO: comment</param>
		public void Bind(ICollection list, string categoryPropertyName, string valuePropertyName)
		{
			XElement
				ptCount = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), list.Count)),
				formatCode = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");
			strCache.RemoveAll();
			numCache.RemoveAll();
			strCache.Add(ptCount);
			numCache.Add(formatCode);
			numCache.Add(ptCount);
			int index = 0;
			XElement pt;
			foreach (var item in list)
			{
				pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index), new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(categoryPropertyName).GetValue(item, null)));
				strCache.Add(pt);
				pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index), new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(valuePropertyName).GetValue(item, null)));
				numCache.Add(pt);
				index++;
			}
		}

		/// <summary>TODO: comment</summary>
		/// <param name="categories">TODO: comment</param>
		/// <param name="values">TODO: comment</param>
		public void Bind(IList categories, IList values)
		{
			if (categories.Count != values.Count) throw new ArgumentException("Categories count must equal to Values count");
			XElement
				ptCount = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), categories.Count)),
				formatCode = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");
			strCache.RemoveAll();
			numCache.RemoveAll();
			strCache.Add(ptCount);
			numCache.Add(formatCode);
			numCache.Add(ptCount);
			XElement pt;
			for (int index = 0; index < categories.Count; index++)
			{
				pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index), new XElement(XName.Get("v", DocX.c.NamespaceName), categories[index].ToString()));
				strCache.Add(pt);
				pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index), new XElement(XName.Get("v", DocX.c.NamespaceName), values[index].ToString()));
				numCache.Add(pt);
			}
		}

	}

	/// <summary>Represents a chart legend, more at http://msdn.microsoft.com/ru-ru/library/cc845123.aspx </summary>
	public class ChartLegend
	{

		/// <summary>Legend xml element</summary>
		internal XElement Xml
		{
			get;
			private set;
		}

		/// <summary>Specifies that other chart elements shall be allowed to overlap this chart element</summary>
		public bool Overlay
		{
			get
			{
				return Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value == "1";
			}
			set
			{
				Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value = GetOverlayValue(value);
			}
		}

		/// <summary>Specifies the possible positions for a legend</summary>
		public ChartLegendPosition Position
		{
			get
			{
				return XElementHelpers.GetValueToEnum<ChartLegendPosition>(Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)), value);
			}
		}

		internal ChartLegend(ChartLegendPosition position, bool overlay)
		{
			Xml = new XElement(XName.Get("legend", DocX.c.NamespaceName), new XElement(XName.Get("legendPos", DocX.c.NamespaceName), new XAttribute("val", XElementHelpers.GetXmlNameFromEnum(position))), new XElement(XName.Get("overlay", DocX.c.NamespaceName), new XAttribute("val", GetOverlayValue(overlay))));
		}

		/// <summary>ECMA-376, page 3840. 21.2.2.132 overlay (Overlay)</summary>
		private string GetOverlayValue(bool overlay)
		{
			return overlay ? "1" : "0";
		}

	}

	/// <summary>This element contains the 2-D line chart series. 21.2.2.97 lineChart (Line Charts).</summary>
	public class LineChart : Chart
	{

		/// <summary>Specifies the kind of grouping for a column, line, or area chart</summary>
		public Grouping Grouping
		{
			get
			{
				return XElementHelpers.GetValueToEnum<Grouping>(ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)), value);
			}
		}

		/// <summary></summary>
		/// <returns></returns>
		protected override XElement CreateChartXml()
		{
			return XElement.Parse(@"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:grouping val=""standard""/></c:lineChart>");
		}

	}

	/// <summary>This element contains the 2-D pie series for this chart. 21.2.2.141 pieChart (Pie Charts).</summary>
	public class PieChart : Chart
	{

		/// <summary>TODO: comment</summary>
		public override bool IsAxisExist
		{
			get
			{
				return false;
			}
		}

		/// <summary>TODO: comment</summary>
		public override short MaxSeriesCount
		{
			get
			{
				return 1;
			}
		}

		/// <summary>TODO: comment</summary>
		/// <returns>TODO: comment</returns>
		protected override XElement CreateChartXml()
		{
			return XElement.Parse(@"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""></c:pieChart>");
		}

	}

	/// <summary>TODO: comment</summary>
	public class Bookmark
	{

		/// <summary>TODO: comment</summary>
		public string Name
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		public Paragraph Paragraph
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		/// <param name="newText">TODO: comment</param>
		public void SetText(string newText)
		{
			Paragraph.ReplaceAtBookmark(newText, Name);
		}

	}

	/// <summary>TODO: comment</summary>
	public class BookmarkCollection : List<Bookmark>
	{

		/// <summary>TODO: comment</summary>
		/// <param name="name">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public Bookmark this[string name]
		{
			get
			{
				return this.FirstOrDefault(bookmark => string.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase));
			}
		}

	}

	/// <summary>Represents a border of a table or table cell. Added by lckuiper @ 20101117.</summary>
	public class Border
	{

		/// <summary>TODO: comment</summary>
		public BorderStyle Tcbs
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		public BorderSize Size
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		public int Space
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		public Color Color
		{
			get;
			set;
		}

		/// <summary>TODO: comment</summary>
		public Border()
		{
			Tcbs = BorderStyle.Tcbs_single;
			Size = BorderSize.one;
			Space = 0;
			Color = Color.Black;
		}

		/// <summary>TODO: comment</summary>
		/// <param name="tcbs">TODO: comment</param>
		/// <param name="size">TODO: comment</param>
		/// <param name="space">TODO: comment</param>
		/// <param name="color">TODO: comment</param>
		public Border(BorderStyle tcbs, BorderSize size, int space, Color color)
		{
			Tcbs = tcbs;
			Size = size;
			Space = space;
			Color = color;
		}

	}

	// TILL HERE +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	/// <summary>TODO: comment</summary>
	public abstract class Container : DocXElement
	{

		/// <summary>TODO: comment</summary>
		public virtual ReadOnlyCollection<Content> Contents
		{
			get
			{
				List<Content> contents = GetContents();
				return contents.AsReadOnly();
			}
		}

		/// <summary>Returns a list of all Paragraphs inside this container</summary>
		/// <example>
		///		<code>
		///			Load a document
		///			using (DocX document = DocX.Load(@"Test.docx"))
		///			{
		///				// All Paragraphs in this document
		///				<![CDATA[ List<Paragraph> ]]> documentParagraphs = document.Paragraphs;
		///				// Make sure this document contains at least one Table
		///				if (document.Tables.Count() > 0)
		///				{
		///					// Get the first Table in this document
		///					Table t = document.Tables[0];
		///					// All Paragraphs in this Table
		///					<![CDATA[ List<Paragraph> ]]> tableParagraphs = t.Paragraphs;
		///					// Make sure this Table contains at least one Row
		///					if (t.Rows.Count() > 0)
		///					{
		///						// Get the first Row in this document
		///						Row r = t.Rows[0];
		///						// All Paragraphs in this Row
		///						<![CDATA[ List<Paragraph> ]]> rowParagraphs = r.Paragraphs;
		///						// Make sure this Row contains at least one Cell
		///						if (r.Cells.Count() > 0)
		///						{
		///							// Get the first Cell in this document
		///							Cell c = r.Cells[0];
		///							// All Paragraphs in this Cell
		///							<![CDATA[ List<Paragraph> ]]> cellParagraphs = c.Paragraphs;
		///						}
		///					}
		///				}
		///				// Save all changes to this document
		///				document.Save();
		///			}
		///			// Release this document from memory
		///		</code>
		/// </example>
		public virtual ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				List<Paragraph> paragraphs = GetParagraphs();
				foreach (var p in paragraphs)
				{
					if ((p.Xml.ElementsAfterSelf().FirstOrDefault() != null) && (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl"))) p.FollowingTable = new Table(Document, p.Xml.ElementsAfterSelf().First());
					p.ParentContainer = GetParentFromXmlName(p.Xml.Ancestors().First().Name.LocalName);
					if (p.IsListItem) GetListItemType(p);
				}
				return paragraphs.AsReadOnly();
			}
		}

		/// <summary>TODO: comment</summary>
		public virtual ReadOnlyCollection<Paragraph> ParagraphsDeepSearch
		{
			get
			{
				List<Paragraph> paragraphs = GetParagraphs(true);
				foreach (var p in paragraphs)
				{
					if ((p.Xml.ElementsAfterSelf().FirstOrDefault() != null) && (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl"))) p.FollowingTable = new Table(Document, p.Xml.ElementsAfterSelf().First());
					p.ParentContainer = GetParentFromXmlName(p.Xml.Ancestors().First().Name.LocalName);
					if (p.IsListItem) GetListItemType(p);
				}
				return paragraphs.AsReadOnly();
			}
		}

		/// <summary>Removes paragraph at specified position</summary>
		/// <param name="index">Index of paragraph to remove</param>
		/// <returns>True if removed</returns>
		public bool RemoveParagraphAt(int index)
		{
			int i = 0;
			foreach (var paragraph in Xml.Descendants(DocX.w + "p"))
			{
				if (i == index)
				{
					paragraph.Remove();
					return true;
				}
				++i;
			}
			return false;
		}

		/// <summary>Removes paragraph</summary>
		/// <param name="p">Paragraph to remove</param>
		/// <returns>True if removed</returns>
		public bool RemoveParagraph(Paragraph p)
		{
			foreach (var paragraph in Xml.Descendants(DocX.w + "p"))
			{
				if (paragraph.Equals(p.Xml))
				{
					paragraph.Remove();
					return true;
				}
			}
			return false;
		}

		/// <summary>TODO: comment</summary>
		public virtual List<Section> Sections
		{
			get
			{
				var allParas = Paragraphs;
				var parasInASection = new List<Paragraph>();
				var sections = new List<Section>();
				foreach (var para in allParas)
				{
					var sectionInPara = para.Xml.Descendants().FirstOrDefault(s => s.Name.LocalName == "sectPr");
					if (sectionInPara == null)
					{
						parasInASection.Add(para);
					}
					else
					{
						parasInASection.Add(para);
						var section = new Section(Document, sectionInPara) { SectionParagraphs = parasInASection };
						sections.Add(section);
						parasInASection = new List<Paragraph>();
					}
				}
				XElement body = Xml.Element(XName.Get("body", DocX.w.NamespaceName));
				XElement baseSectionXml = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
				var baseSection = new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection };
				sections.Add(baseSection);
				return sections;
			}
		}

		private void GetListItemType(Paragraph p)
		{
			var ilvlNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(el => el.Name.LocalName == "ilvl");
			var ilvlValue = ilvlNode.Attribute(DocX.w + "val").Value;
			var numIdNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(el => el.Name.LocalName == "numId");
			var numIdValue = numIdNode.Attribute(DocX.w + "val").Value;
			// find num node in numbering
			var numNodes = Document.numbering.Descendants().Where(n => n.Name.LocalName == "num");
			XElement numNode = numNodes.FirstOrDefault(node => node.Attribute(DocX.w + "numId").Value.Equals(numIdValue));
			if (numNode != null)
			{
				// Get abstractNumId node and its value from numNode
				var abstractNumIdNode = numNode.Descendants().First(n => n.Name.LocalName == "abstractNumId");
				var abstractNumNodeValue = abstractNumIdNode.Attribute(DocX.w + "val").Value;
				var abstractNumNodes = Document.numbering.Descendants().Where(n => n.Name.LocalName == "abstractNum");
				XElement abstractNumNode = abstractNumNodes.FirstOrDefault(node => node.Attribute(DocX.w + "abstractNumId").Value.Equals(abstractNumNodeValue));
				// Find lvl node
				var lvlNodes = abstractNumNode.Descendants().Where(n => n.Name.LocalName == "lvl");
				XElement lvlNode = null;
				foreach (XElement node in lvlNodes)
				{
					if (node.Attribute(DocX.w + "ilvl").Value.Equals(ilvlValue))
					{
						lvlNode = node;
						break;
					}
				}
				var numFmtNode = lvlNode.Descendants().First(n => n.Name.LocalName == "numFmt");
				p.ListItemType = GetListItemType(numFmtNode.Attribute(DocX.w + "val").Value);
			}
		}

		/// <summary>TODO: comment</summary>
		public ContainerType ParentContainer;

		internal List<Content> GetContents(bool deepSearch = false)
		{
			// Need some memory that can be updated by the recursive search
			List<Content> contents = new List<Content>();
			foreach (XElement e in Xml.Descendants(XName.Get("sdt", DocX.w.NamespaceName)))
			{
				Content content = new Content(Document, e, 0);
				XElement el = e.Elements(XName.Get("sdtPr", DocX.w.NamespaceName)).First();
				content.Name = GetAttribute(el, "alias", "val");
				content.Tag = GetAttribute(el, "tag", "val");
				contents.Add(content);
			}
			return contents;
		}

		private string GetAttribute(XElement e, string localName, string attributeName)
		{
			string val = string.Empty;
			try
			{
				val = e.Elements(XName.Get(localName, DocX.w.NamespaceName)).Attributes(XName.Get(attributeName, DocX.w.NamespaceName)).FirstOrDefault().Value;
			}
			catch (Exception)
			{
				val = "Missing";
			}
			return val;
		}

		internal List<Paragraph> GetParagraphs(bool deepSearch = false)
		{
			// Need some memory that can be updated by the recursive search
			int index = 0;
			List<Paragraph> paragraphs = new List<Paragraph>();
			foreach (XElement e in Xml.Descendants(XName.Get("p", DocX.w.NamespaceName)))
			{
				index += HelperFunctions.GetText(e).Length;
				Paragraph paragraph = new Paragraph(Document, e, index);
				paragraphs.Add(paragraph);
			}
			return paragraphs;
		}

		internal void GetParagraphsRecursive(XElement Xml, ref int index, ref List<Paragraph> paragraphs, bool deepSearch = false)
		{
			var keepSearching = true;
			if (Xml.Name.LocalName == "p")
			{
				paragraphs.Add(new Paragraph(Document, Xml, index));
				index += HelperFunctions.GetText(Xml).Length;
				if (!deepSearch) keepSearching = false;
			}
			if (keepSearching && Xml.HasElements) foreach (XElement e in Xml.Elements()) GetParagraphsRecursive(e, ref index, ref paragraphs, deepSearch);
		}

		/// <summary>TODO: comment</summary>
		public virtual List<Table> Tables
		{
			get
			{
				List<Table> tables = (from t in Xml.Descendants(DocX.w + "tbl") select new Table(Document, t)).ToList();
				return tables;
			}
		}

		/// <summary>TODO: comment</summary>
		public virtual List<List> Lists
		{
			get
			{
				var lists = new List<List>();
				var list = new List(Document, Xml);
				foreach (var paragraph in Paragraphs)
				{
					if (paragraph.IsListItem)
					{
						if (list.CanAddListItem(paragraph))
						{
							list.AddItem(paragraph);
						}
						else
						{
							lists.Add(list);
							list = new List(Document, Xml);
							list.AddItem(paragraph);
						}
					}
				}
				lists.Add(list);
				return lists;
			}
		}

		/// <summary>TODO: comment</summary>
		public virtual List<Hyperlink> Hyperlinks
		{
			get
			{
				List<Hyperlink> hyperlinks = new List<Hyperlink>();
				foreach (Paragraph p in Paragraphs) hyperlinks.AddRange(p.Hyperlinks);
				return hyperlinks;
			}
		}

		/// <summary>TODO: comment</summary>
		public virtual List<Picture> Pictures
		{
			get
			{
				List<Picture> pictures = new List<Picture>();
				foreach (Paragraph p in Paragraphs) pictures.AddRange(p.Pictures);
				return pictures;
			}
		}

		/// <summary>Sets the Direction of content</summary>
		/// <param name="direction">Direction either LeftToRight or RightToLeft</param>
		/// <example>
		///		Set the Direction of content in a Paragraph to RightToLeft
		///		<code>
		///			// Load a document
		///			using (DocX document = DocX.Load(@"Test.docx"))
		///			{
		///				// Get the first Paragraph from this document
		///				Paragraph p = document.InsertParagraph();
		///				// Set the Direction of this Paragraph
		///				p.Direction = Direction.RightToLeft;
		///				// Make sure the document contains at lest one Table
		///				if (document.Tables.Count() > 0)
		///				{
		///					// Get the first Table from this document
		///					Table t = document.Tables[0];
		///					// Set the direction of the entire Table
		///					// Note: The same function is available at the Row and Cell level
		///					t.SetDirection(Direction.RightToLeft);
		///				}
		///				// Save all changes to this document
		///				document.Save();
		///			}
		///			// Release this document from memory
		///		</code>
		/// </example>
		public virtual void SetDirection(Direction direction)
		{
			foreach (Paragraph p in Paragraphs) p.Direction = direction;
		}

		/// <summary>TODO: comment</summary>
		/// <param name="str">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual List<int> FindAll(string str)
		{
			return FindAll(str, RegexOptions.None);
		}

		/// <summary>TODO: comment</summary>
		/// <param name="str">TODO: comment</param>
		/// <param name="options">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual List<int> FindAll(string str, RegexOptions options)
		{
			List<int> list = new List<int>();
			foreach (Paragraph p in Paragraphs)
			{
				List<int> indexes = p.FindAll(str, options);
				for (int i = 0; i < indexes.Count(); i++) indexes[0] += p.startIndex;
				list.AddRange(indexes);
			}
			return list;
		}

		/// <summary>Find all unique instances of the given Regex Pattern, returning the list of the unique strings found</summary>
		/// <param name="pattern">TODO: comment</param>
		/// <param name="options">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual List<string> FindUniqueByPattern(string pattern, RegexOptions options)
		{
			List<string> rawResults = new List<string>();
			foreach (Paragraph p in Paragraphs)
			{
				// accumulate the search results from all paragraphs
				List<string> partials = p.FindAllByPattern(pattern, options);
				rawResults.AddRange(partials);
			}
			// this dictionary is used to collect results and test for uniqueness
			Dictionary<string, int> uniqueResults = new Dictionary<string, int>();
			foreach (string currValue in rawResults)
			{
				// if the dictionary doesn't have it, add it
				if (!uniqueResults.ContainsKey(currValue)) uniqueResults.Add(currValue, 0);
			}
			// return the unique list of results
			return uniqueResults.Keys.ToList();
		}

		/// <summary>TODO: comment</summary>
		/// <param name="searchValue">TODO: comment</param>
		/// <param name="newValue">TODO: comment</param>
		/// <param name="trackChanges">TODO: comment</param>
		/// <param name="options">TODO: comment</param>
		/// <param name="newFormatting">TODO: comment</param>
		/// <param name="matchFormatting">TODO: comment</param>
		/// <param name="formattingOptions">TODO: comment</param>
		/// <param name="escapeRegEx">TODO: comment</param>
		/// <param name="useRegExSubstitutions">TODO: comment</param>
		public virtual void ReplaceText(string searchValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
		{
			if (string.IsNullOrEmpty(searchValue)) throw new ArgumentException("oldValue cannot be null or empty", "searchValue");
			if (newValue == null) throw new ArgumentException("newValue cannot be null or empty", "newValue");
			// ReplaceText in Headers of the document
			var headerList = new List<Header>
			{
				Document.Headers.first,
				Document.Headers.even,
				Document.Headers.odd
			};
			foreach (var header in headerList) if (header != null) foreach (var paragraph in header.Paragraphs) paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
			// ReplaceText int main body of document
			foreach (var paragraph in Paragraphs) paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
			// ReplaceText in Footers of the document
			var footerList = new List<Footer>
			{
				Document.Footers.first,
				Document.Footers.even,
				Document.Footers.odd
			};
			foreach (var footer in footerList) if (footer != null) foreach (var paragraph in footer.Paragraphs) paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
		}

		/// <summary>TODO: comment</summary>
		/// <param name="searchValue">Value to find</param>
		/// <param name="regexMatchHandler">A Func that accepts the matching regex search group value and passes it to this to return the replacement string</param>
		/// <param name="trackChanges">Enable trackchanges</param>
		/// <param name="options">Regex options</param>
		/// <param name="newFormatting">TODO: comment</param>
		/// <param name="matchFormatting">TODO: comment</param>
		/// <param name="formattingOptions">TODO: comment</param>
		public virtual void ReplaceText(string searchValue, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch)
		{
			if (string.IsNullOrEmpty(searchValue)) throw new ArgumentException("oldValue cannot be null or empty", "searchValue");
			if (regexMatchHandler == null) throw new ArgumentException("regexMatchHandler cannot be null", "regexMatchHandler");
			// ReplaceText in Headers/Footers of the document
			var containerList = new List<IParagraphContainer>
			{
				Document.Headers.first, Document.Headers.even, Document.Headers.odd,
				Document.Footers.first, Document.Footers.even, Document.Footers.odd
			};
			foreach (var container in containerList) if (container != null) foreach (var paragraph in container.Paragraphs) paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
			// ReplaceText int main body of document
			foreach (var paragraph in Paragraphs) paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
		}

		/// <summary>Removes all items with required formatting</summary>
		/// <returns>Numer of texts removed</returns>
		public int RemoveTextInGivenFormat(Formatting matchFormatting, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
		{
			var deletedCount = 0;
			foreach (var x in Xml.Elements()) deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo);
			return deletedCount;
		}

		internal int RemoveTextWithFormatRecursive(XElement element, Formatting matchFormatting, MatchFormattingOptions fo)
		{
			var deletedCount = 0;
			foreach (var x in element.Elements())
			{
				if ("rPr".Equals(x.Name.LocalName))
				{
					if (HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, x, fo))
					{
						x.Parent.Remove();
						++deletedCount;
					}
				}
				deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo);
			}
			return deletedCount;
		}

		/// <summary>TODO: comment</summary>
		/// <param name="toInsert">TODO: comment</param>
		/// <param name="bookmarkName">TODO: comment</param>
		public virtual void InsertAtBookmark(string toInsert, string bookmarkName)
		{
			if (bookmarkName.IsNullOrWhiteSpace()) throw new ArgumentException("bookmark cannot be null or empty", "bookmarkName");
			var headerCollection = Document.Headers;
			var headers = new List<Header>
			{
				headerCollection.first,
				headerCollection.even,
				headerCollection.odd
			};
			foreach (var header in headers.Where(x => x != null)) foreach (var paragraph in header.Paragraphs) paragraph.InsertAtBookmark(toInsert, bookmarkName);
			foreach (var paragraph in Paragraphs) paragraph.InsertAtBookmark(toInsert, bookmarkName);
			var footerCollection = Document.Footers;
			var footers = new List<Footer>
			{
				footerCollection.first,
				footerCollection.even,
				footerCollection.odd
			};
			foreach (var footer in footers.Where(x => x != null)) foreach (var paragraph in footer.Paragraphs) paragraph.InsertAtBookmark(toInsert, bookmarkName);
		}

		/// <summary>TODO: comment</summary>
		/// <param name="bookmarkNames">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public string[] ValidateBookmarks(params string[] bookmarkNames)
		{
			var headers = new[] { Document.Headers.first, Document.Headers.even, Document.Headers.odd }.Where(h => h != null).ToList();
			var footers = new[] { Document.Footers.first, Document.Footers.even, Document.Footers.odd }.Where(f => f != null).ToList();
			var nonMatching = new List<string>();
			foreach (var bookmarkName in bookmarkNames)
			{
				if (headers.SelectMany(h => h.Paragraphs).Any(p => p.ValidateBookmark(bookmarkName))) return new string[0];
				if (footers.SelectMany(h => h.Paragraphs).Any(p => p.ValidateBookmark(bookmarkName))) return new string[0];
				if (Paragraphs.Any(p => p.ValidateBookmark(bookmarkName))) return new string[0];
				nonMatching.Add(bookmarkName);
			}
			return nonMatching.ToArray();
		}

		/// <summary>TODO: comment</summary>
		/// <param name="index">TODO: comment</param>
		/// <param name="text">TODO: comment</param>
		/// <param name="trackChanges">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			return InsertParagraph(index, text, trackChanges, null);
		}

		/// <summary>TODO: comment</summary>
		/// <returns>TODO: comment</returns>
		public virtual Paragraph InsertParagraph()
		{
			return InsertParagraph(string.Empty, false);
		}

		/// <summary>TODO: comment</summary>
		/// <param name="index">TODO: comment</param>
		/// <param name="p">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual Paragraph InsertParagraph(int index, Paragraph p)
		{
			XElement newXElement = new XElement(p.Xml);
			p.Xml = newXElement;
			Paragraph paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);
			if (paragraph == null) Xml.Add(p.Xml);
			else
			{
				XElement[] split = HelperFunctions.SplitParagraph(paragraph, index - paragraph.startIndex);
				paragraph.Xml.ReplaceWith(split[0], newXElement, split[1]);
			}
			GetParent(p);
			return p;
		}

		/// <summary>TODO: comment</summary>
		/// <param name="p">TODO: comment</param>
		/// <returns>TODO: comment</returns>
		public virtual Paragraph InsertParagraph(Paragraph p)
		{
			XDocument style_document;
			if (p.styles.Count() > 0)
			{
				Uri style_package_uri = new Uri("/word/styles.xml", UriKind.Relative);
				if (!Document.package.PartExists(style_package_uri))
				{
					PackagePart style_package = Document.package.CreatePart(style_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
					using (TextWriter tw = new StreamWriter(new PackagePartStream(style_package.GetStream())))
					{
						style_document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("styles", DocX.w.NamespaceName)));
						style_document.Save(tw);
					}
				}
				PackagePart styles_document = Document.package.GetPart(style_package_uri);
				using (TextReader tr = new StreamReader(styles_document.GetStream()))
				{
					style_document = XDocument.Load(tr);
					XElement styles_element = style_document.Element(XName.Get("styles", DocX.w.NamespaceName));
					var ids = from d in styles_element.Descendants(XName.Get("style", DocX.w.NamespaceName)) let a = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName)) where a != null select a.Value;
					foreach (XElement style in p.styles)
					{
						// If styles_element does not contain this element, then add it
						if (!ids.Contains(style.Attribute(XName.Get("styleId", DocX.w.NamespaceName)).Value)) styles_element.Add(style);
					}
				}
				using (TextWriter tw = new StreamWriter(new PackagePartStream(styles_document.GetStream()))) style_document.Save(tw);
			}
			XElement newXElement = new XElement(p.Xml);
			Xml.Add(newXElement);
			int index = 0;
			if (Document.paragraphLookup.Keys.Count() > 0)
			{
				index = Document.paragraphLookup.Last().Key;
				if (Document.paragraphLookup.Last().Value.Text.Length == 0) index++;
				else index += Document.paragraphLookup.Last().Value.Text.Length;
			}
			Paragraph newParagraph = new Paragraph(Document, newXElement, index);
			Document.paragraphLookup.Add(index, newParagraph);
			GetParent(newParagraph);
			return newParagraph;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph newParagraph = new Paragraph(Document, new XElement(DocX.w + "p"), index);
			newParagraph.InsertText(0, text, trackChanges, formatting);
			Paragraph firstPar = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);
			if (firstPar != null)
			{
				var splitindex = index - firstPar.startIndex;
				if (splitindex <= 0)
				{
					firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml);
				}
				else
				{
					XElement[] splitParagraph = HelperFunctions.SplitParagraph(firstPar, splitindex);
					firstPar.Xml.ReplaceWith
					(
						splitParagraph[0],
						newParagraph.Xml,
						splitParagraph[1]
					);
				}
			}
			else Xml.Add(newParagraph);
			GetParent(newParagraph);
			return newParagraph;
		}

		private ContainerType GetParentFromXmlName(string xmlName)
		{
			ContainerType parent;
			switch (xmlName)
			{
				case "body":
					parent = ContainerType.Body;
					break;
				case "p":
					parent = ContainerType.Paragraph;
					break;
				case "tbl":
					parent = ContainerType.Table;
					break;
				case "sectPr":
					parent = ContainerType.Section;
					break;
				case "tc":
					parent = ContainerType.Cell;
					break;
				default:
					parent = ContainerType.None;
					break;
			}
			return parent;
		}

		private void GetParent(Paragraph newParagraph)
		{
			var containerType = GetType();
			switch (containerType.Name)
			{
				case "Body":
					newParagraph.ParentContainer = ContainerType.Body;
					break;
				case "Table":
					newParagraph.ParentContainer = ContainerType.Table;
					break;
				case "TOC":
					newParagraph.ParentContainer = ContainerType.TOC;
					break;
				case "Section":
					newParagraph.ParentContainer = ContainerType.Section;
					break;
				case "Cell":
					newParagraph.ParentContainer = ContainerType.Cell;
					break;
				case "Header":
					newParagraph.ParentContainer = ContainerType.Header;
					break;
				case "Footer":
					newParagraph.ParentContainer = ContainerType.Footer;
					break;
				case "Paragraph":
					newParagraph.ParentContainer = ContainerType.Paragraph;
					break;
			}
		}

		private ListItemType GetListItemType(string styleName)
		{
			ListItemType listItemType;
			switch (styleName)
			{
				case "bullet":
					listItemType = ListItemType.Bulleted;
					break;
				default:
					listItemType = ListItemType.Numbered;
					break;
			}
			return listItemType;
		}

		/// <summary></summary>
		public virtual void InsertSection()
		{
			InsertSection(false);
		}

		/// <summary></summary>
		/// <param name="trackChanges"></param>
		public virtual void InsertSection(bool trackChanges)
		{
			var newParagraphSection = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName), new XElement(XName.Get("type", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", "continuous"))))
			);
			if (trackChanges) newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);
			Xml.Add(newParagraphSection);
		}

		/// <summary></summary>
		/// <param name="trackChanges"></param>
		public virtual void InsertSectionPageBreak(bool trackChanges = false)
		{
			var newParagraphSection = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName)))
			);
			if (trackChanges) newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);
			Xml.Add(newParagraphSection);
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraph(string text)
		{
			return InsertParagraph(text, false, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraph(string text, bool trackChanges)
		{
			return InsertParagraph(text, trackChanges, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			XElement newParagraph = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
			);
			if (trackChanges) newParagraph = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraph);
			Xml.Add(newParagraph);
			var paragraphAdded = new Paragraph(Document, newParagraph, 0);
			if (this is Cell)
			{
				var cell = this as Cell;
				paragraphAdded.PackagePart = cell.mainPart;
			}
			else if (this is DocX)
			{
				paragraphAdded.PackagePart = Document.mainPart;
			}
			else if (this is Footer)
			{
				var f = this as Footer;
				paragraphAdded.mainPart = f.mainPart;
			}
			else if (this is Header)
			{
				var h = this as Header;
				paragraphAdded.mainPart = h.mainPart;
			}
			else
			{
				Console.WriteLine("No idea what we are {0}", this);
				paragraphAdded.PackagePart = Document.mainPart;
			}
			GetParent(paragraphAdded);
			return paragraphAdded;
		}

		/// <summary></summary>
		/// <param name="equation"></param>
		/// <returns></returns>
		public virtual Paragraph InsertEquation(string equation)
		{
			Paragraph p = InsertParagraph();
			p.AppendEquation(equation);
			return p;
		}

		/// <summary></summary>
		/// <param name="bookmarkName"></param>
		/// <returns></returns>
		public virtual Paragraph InsertBookmark(string bookmarkName)
		{
			var p = InsertParagraph();
			p.AppendBookmark(bookmarkName);
			return p;
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public virtual Table InsertTable(int rowCount, int columnCount) //Dmitchern, changed to virtual, and overrided in Table.Cell
		{
			XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
			Xml.Add(newTable);
			return new Table(Document, newTable) { mainPart = mainPart };
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public Table InsertTable(int index, int rowCount, int columnCount)
		{
			XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
			Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);
			if (p == null) Xml.Elements().First().AddFirst(newTable);
			else
			{
				XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
				p.Xml.ReplaceWith(split[0], newTable, split[1]);
			}
			return new Table(Document, newTable) { mainPart = mainPart };
		}

		/// <summary></summary>
		/// <param name="t"></param>
		/// <returns></returns>
		public Table InsertTable(Table t)
		{
			XElement newXElement = new XElement(t.Xml);
			Xml.Add(newXElement);
			Table newTable = new Table(Document, newXElement)
			{
				mainPart = mainPart,
				Design = t.Design
			};
			return newTable;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="t"></param>
		/// <returns></returns>
		public Table InsertTable(int index, Table t)
		{
			Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);
			XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
			XElement newXElement = new XElement(t.Xml);
			p.Xml.ReplaceWith(split[0], newXElement, split[1]);
			Table newTable = new Table(Document, newXElement)
			{
				mainPart = mainPart,
				Design = t.Design
			};
			return newTable;
		}

		internal Container(DocX document, XElement xml) : base(document, xml)
		{
		}

		/// <summary></summary>
		/// <param name="list"></param>
		/// <returns></returns>
		public List InsertList(List list)
		{
			foreach (var item in list.Items) Xml.Add(item.Xml);
			return list;
		}

		/// <summary></summary>
		/// <param name="list"></param>
		/// <param name="fontSize"></param>
		/// <returns></returns>
		public List InsertList(List list, double fontSize)
		{
			foreach (var item in list.Items)
			{
				item.FontSize(fontSize);
				Xml.Add(item.Xml);
			}
			return list;
		}

		/// <summary></summary>
		/// <param name="list"></param>
		/// <param name="fontFamily"></param>
		/// <param name="fontSize"></param>
		/// <returns></returns>
		public List InsertList(List list, Font fontFamily, double fontSize)
		{
			foreach (var item in list.Items)
			{
				item.Font(fontFamily);
				item.FontSize(fontSize);
				Xml.Add(item.Xml);
			}
			return list;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="list"></param>
		/// <returns></returns>
		public List InsertList(int index, List list)
		{
			Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);
			XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
			var elements = new List<XElement> { split[0] };
			elements.AddRange(list.Items.Select(i => new XElement(i.Xml)));
			elements.Add(split[1]);
			p.Xml.ReplaceWith(elements.ToArray());
			return list;
		}

	}

	/// <summary></summary>
	public class Content : InsertBeforeOrAfter
	{

		/// <summary></summary>
		public string Name { get; set; }

		/// <summary></summary>
		public string Tag { get; set; }

		/// <summary></summary>
		public string Text { get; set; }

		/// <summary></summary>
		public List<Content> Sections { get; set; }

		/// <summary></summary>
		public ContainerType ParentContainer;

		internal Content(DocX document, XElement xml, int startIndex) : base(document, xml)
		{
		}

		/// <summary></summary>
		/// <param name="newText"></param>
		public void SetText(string newText)
		{
			string x = this.Tag;
			XElement el = Xml;
			Xml.Descendants(XName.Get("t", DocX.w.NamespaceName)).First().Value = newText;
		}

	}

	/// <summary></summary>
	public class ContentCollection : List<Content>
	{

		/// <summary></summary>
		/// <param name="name"></param>
		/// <returns></returns>
		public Content this[string name]
		{
			get
			{
				return this.FirstOrDefault(content => string.Equals(content.Name, name, StringComparison.CurrentCultureIgnoreCase));
			}
		}

	}

	/// <summary></summary>
	public class CustomProperty
	{

		private string name;

		private object value;

		private string type;

		/// <summary>The name of this CustomProperty</summary>
		public string Name
		{
			get
			{
				return name;
			}
		}

		/// <summary>The value of this CustomProperty</summary>
		public object Value
		{
			get
			{
				return value;
			}
		}

		internal string Type
		{
			get
			{
				return type;
			}
		}

		internal CustomProperty(string name, string type, string value)
		{
			object realValue;
			switch (type)
			{
				case "lpwstr":
					{
						realValue = value;
						break;
					}
				case "i4":
					{
						realValue = int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
						break;
					}
				case "r8":
					{
						realValue = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
						break;
					}
				case "filetime":
					{
						realValue = DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
						break;
					}
				case "bool":
					{
						realValue = bool.Parse(value);
						break;
					}
				default: throw new Exception();
			}
			this.name = name;
			this.type = type;
			this.value = realValue;
		}

		private CustomProperty(string name, string type, object value)
		{
			this.name = name;
			this.type = type;
			this.value = value;
		}

		/// <summary>Create a new CustomProperty to hold a string</summary>
		/// <param name="name">The name of this CustomProperty.</param>
		/// <param name="value">The value of this CustomProperty.</param>
		public CustomProperty(string name, string value) : this(name, "lpwstr", value as object)
		{
		}

		/// <summary>Create a new CustomProperty to hold an int</summary>
		/// <param name="name">The name of this CustomProperty.</param>
		/// <param name="value">The value of this CustomProperty.</param>
		public CustomProperty(string name, int value) : this(name, "i4", value as object)
		{
		}

		/// <summary>Create a new CustomProperty to hold a double</summary>
		/// <param name="name">The name of this CustomProperty.</param>
		/// <param name="value">The value of this CustomProperty.</param>
		public CustomProperty(string name, double value) : this(name, "r8", value as object)
		{
		}

		/// <summary>Create a new CustomProperty to hold a DateTime</summary>
		/// <param name="name">The name of this CustomProperty.</param>
		/// <param name="value">The value of this CustomProperty.</param>
		public CustomProperty(string name, DateTime value) : this(name, "filetime", value.ToUniversalTime() as object)
		{
		}

		/// <summary>Create a new CustomProperty to hold a bool</summary>
		/// <param name="name">The name of this CustomProperty.</param>
		/// <param name="value">The value of this CustomProperty.</param>
		public CustomProperty(string name, bool value) : this(name, "bool", value as object)
		{
		}

	}

	/// <summary>Represents a field of type document property. This field displays the value stored in a custom property.</summary>
	public class DocProperty : DocXElement
	{

		internal Regex extractName = new Regex(@"DOCPROPERTY  (?<name>.*)  ");

		private string name;

		/// <summary>The custom property to display</summary>
		public string Name
		{
			get
			{
				return name;
			}
		}

		internal DocProperty(DocX document, XElement xml) : base(document, xml)
		{
			string instr = Xml.Attribute(XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Value;
			name = extractName.Match(instr.Trim()).Groups["name"].Value;
		}

	}

	/// <summary>Represents a document</summary>
	public class DocX : Container, IDisposable
	{

		static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		static internal XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

		static internal XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		static internal XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";

		static internal XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

		static internal XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

		static internal XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

		static internal XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

		static internal XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";

		static internal XNamespace v = "urn:schemas-microsoft-com:vml";

		internal static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

		internal float getMarginAttribute(XName name)
		{
			XElement body = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XElement sectPr = body.Element(XName.Get("sectPr", w.NamespaceName));
			if (sectPr != null)
			{
				XElement pgMar = sectPr.Element(XName.Get("pgMar", w.NamespaceName));
				if (pgMar != null)
				{
					XAttribute top = pgMar.Attribute(name);
					if (top != null)
					{
						float f;
						if (float.TryParse(top.Value, out f)) return (int)(f / 20.0f);
					}
				}
			}
			return 0;
		}

		internal void setMarginAttribute(XName xName, float value)
		{
			XElement body = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XElement sectPr = body.Element(XName.Get("sectPr", w.NamespaceName));
			if (sectPr != null)
			{
				XElement pgMar = sectPr.Element(XName.Get("pgMar", w.NamespaceName));
				if (pgMar != null)
				{
					XAttribute top = pgMar.Attribute(xName);
					if (top != null)
					{
						top.SetValue(value * 20);
					}
				}
			}
		}

		/// <summary></summary>
		public BookmarkCollection Bookmarks
		{
			get
			{
				BookmarkCollection bookmarks = new BookmarkCollection();
				foreach (Paragraph paragraph in Paragraphs) bookmarks.AddRange(paragraph.GetBookmarks());
				return bookmarks;
			}
		}

		/// <summary>Top margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.</summary>
		public float MarginTop
		{
			get
			{
				return getMarginAttribute(XName.Get("top", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("top", w.NamespaceName), value);
			}
		}

		/// <summary>Bottom margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.</summary>
		public float MarginBottom
		{
			get
			{
				return getMarginAttribute(XName.Get("bottom", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("bottom", w.NamespaceName), value);
			}
		}

		/// <summary>Left margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.</summary>
		public float MarginLeft
		{
			get
			{
				return getMarginAttribute(XName.Get("left", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("left", w.NamespaceName), value);
			}
		}

		/// <summary>Right margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.</summary>
		public float MarginRight
		{
			get
			{
				return getMarginAttribute(XName.Get("right", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("right", w.NamespaceName), value);
			}
		}

		/// <summary>Page width value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.</summary>
		public float PageWidth
		{
			get
			{
				XElement body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
				XElement sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
				if (sectPr != null)
				{
					XElement pgSz = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName));

					if (pgSz != null)
					{
						XAttribute w = pgSz.Attribute(XName.Get("w", DocX.w.NamespaceName));
						if (w != null)
						{
							float f;
							if (float.TryParse(w.Value, out f))
								return (int)(f / 20.0f);
						}
					}
				}

				return (12240.0f / 20.0f);
			}

			set
			{
				XElement body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));

				if (body != null)
				{
					XElement sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));

					if (sectPr != null)
					{
						XElement pgSz = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName));

						if (pgSz != null)
						{
							pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), value * 20);
						}
					}
				}
			}
		}

		/// <summary>
		/// Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float PageHeight
		{
			get
			{
				XElement body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
				XElement sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
				if (sectPr != null)
				{
					XElement pgSz = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName));

					if (pgSz != null)
					{
						XAttribute w = pgSz.Attribute(XName.Get("h", DocX.w.NamespaceName));
						if (w != null)
						{
							float f;
							if (float.TryParse(w.Value, out f))
								return (int)(f / 20.0f);
						}
					}
				}

				return (15840.0f / 20.0f);
			}

			set
			{
				XElement body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));

				if (body != null)
				{
					XElement sectPr = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));

					if (sectPr != null)
					{
						XElement pgSz = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName));

						if (pgSz != null)
						{
							pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), value * 20);
						}
					}
				}
			}
		}
		/// <summary>
		/// Returns true if any editing restrictions are imposed on this document.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     if(document.isProtected)
		///         Console.WriteLine("Protected");
		///     else
		///         Console.WriteLine("Not protected");
		///         
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AddProtection(EditRestrictions)"/>
		/// <seealso cref="RemoveProtection"/>
		/// <seealso cref="GetProtectionType"/>
		public bool isProtected
		{
			get
			{
				return settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).Count() > 0;
			}
		}

		/// <summary>
		/// Returns the type of editing protection imposed on this document.
		/// </summary>
		/// <returns>The type of editing protection imposed on this document.</returns>
		/// <example>
		/// <code>
		/// Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Make sure the document is protected before checking the protection type.
		///     if (document.isProtected)
		///     {
		///         EditRestrictions protection = document.GetProtectionType();
		///         Console.WriteLine("Document is protected using " + protection.ToString());
		///     }
		///
		///     else
		///         Console.WriteLine("Document is not protected.");
		///
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AddProtection(EditRestrictions)"/>
		/// <seealso cref="RemoveProtection"/>
		/// <seealso cref="isProtected"/>
		public EditRestrictions GetProtectionType()
		{
			if (isProtected)
			{
				XElement documentProtection = settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).FirstOrDefault();
				string edit_type = documentProtection.Attribute(XName.Get("edit", DocX.w.NamespaceName)).Value;
				return (EditRestrictions)Enum.Parse(typeof(EditRestrictions), edit_type);
			}

			return EditRestrictions.none;
		}

		/// <summary>
		/// Add editing protection to this document. 
		/// </summary>
		/// <param name="er">The type of protection to add to this document.</param>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Allow no editing, only the adding of comment.
		///     document.AddProtection(EditRestrictions.comments);
		///     
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="RemoveProtection"/>
		/// <seealso cref="GetProtectionType"/>
		/// <seealso cref="isProtected"/>
		public void AddProtection(EditRestrictions er)
		{
			// Call remove protection before adding a new protection element.
			RemoveProtection();

			if (er == EditRestrictions.none)
				return;

			XElement documentProtection = new XElement(XName.Get("documentProtection", DocX.w.NamespaceName));
			documentProtection.Add(new XAttribute(XName.Get("edit", DocX.w.NamespaceName), er.ToString()));
			documentProtection.Add(new XAttribute(XName.Get("enforcement", DocX.w.NamespaceName), "1"));

			settings.Root.AddFirst(documentProtection);
		}

		/// <summary></summary>
		/// <param name="er"></param>
		/// <param name="strPassword"></param>
		public void AddProtection(EditRestrictions er, string strPassword)
		{
			// http://blogs.msdn.com/b/vsod/archive/2010/04/05/how-to-set-the-editing-restrictions-in-word-using-open-xml-sdk-2-0.aspx
			// Call remove protection before adding a new protection element.
			RemoveProtection();

			if (er == EditRestrictions.none)
				return;

			XElement documentProtection = new XElement(XName.Get("documentProtection", DocX.w.NamespaceName));
			documentProtection.Add(new XAttribute(XName.Get("edit", DocX.w.NamespaceName), er.ToString()));
			documentProtection.Add(new XAttribute(XName.Get("enforcement", DocX.w.NamespaceName), "1"));

			int[] InitialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
			int[,] EncryptionMatrix = new int[15, 7]
			{

        /* char 1  */ {0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09},
        /* char 2  */ {0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF},
        /* char 3  */ {0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0},
        /* char 4  */ {0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40},
        /* char 5  */ {0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5},
        /* char 6  */ {0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A},
        /* char 7  */ {0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9},
        /* char 8  */ {0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0},
        /* char 9  */ {0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC},
        /* char 10 */ {0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10},
        /* char 11 */ {0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168},
        /* char 12 */ {0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C},
        /* char 13 */ {0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD},
        /* char 14 */ {0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC},
        /* char 15 */ {0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4}
			};

			// Generate the Salt
			byte[] arrSalt = new byte[16];
			RandomNumberGenerator rand = new RNGCryptoServiceProvider();
			rand.GetNonZeroBytes(arrSalt);

			//Array to hold Key Values
			byte[] generatedKey = new byte[4];

			//Maximum length of the password is 15 chars.
			int intMaxPasswordLength = 15;

			if (!String.IsNullOrEmpty(strPassword))
			{
				strPassword = strPassword.Substring(0, Math.Min(strPassword.Length, intMaxPasswordLength));

				byte[] arrByteChars = new byte[strPassword.Length];

				for (int intLoop = 0; intLoop < strPassword.Length; intLoop++)
				{
					int intTemp = Convert.ToInt32(strPassword[intLoop]);
					arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
					if (arrByteChars[intLoop] == 0)
						arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
				}

				int intHighOrderWord = InitialCodeArray[arrByteChars.Length - 1];

				for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
				{
					int tmp = intMaxPasswordLength - arrByteChars.Length + intLoop;
					for (int intBit = 0; intBit < 7; intBit++)
					{
						if ((arrByteChars[intLoop] & (0x0001 << intBit)) != 0)
						{
							intHighOrderWord ^= EncryptionMatrix[tmp, intBit];
						}
					}
				}

				int intLowOrderWord = 0;

				// For each character in the strPassword, going backwards
				for (int intLoopChar = arrByteChars.Length - 1; intLoopChar >= 0; intLoopChar--)
				{
					intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars[intLoopChar];
				}

				intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars.Length ^ 0xCE4B;

				// Combine the Low and High Order Word
				int intCombinedkey = (intHighOrderWord << 16) + intLowOrderWord;

				// The byte order of the result shall be reversed [Example: 0x64CEED7E becomes 7EEDCE64. end example],
				// and that value shall be hashed as defined by the attribute values.

				for (int intTemp = 0; intTemp < 4; intTemp++)
				{
					generatedKey[intTemp] = Convert.ToByte(((uint)(intCombinedkey & (0x000000FF << (intTemp * 8)))) >> (intTemp * 8));
				}
			}

			StringBuilder sb = new StringBuilder();
			for (int intTemp = 0; intTemp < 4; intTemp++)
			{
				sb.Append(Convert.ToString(generatedKey[intTemp], 16));
			}
			generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());

			byte[] tmpArray1 = generatedKey;
			byte[] tmpArray2 = arrSalt;
			byte[] tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
			Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
			Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
			generatedKey = tempKey;


			int iterations = 100000;

			HashAlgorithm sha1 = new SHA1Managed();
			generatedKey = sha1.ComputeHash(generatedKey);
			byte[] iterator = new byte[4];
			for (int intTmp = 0; intTmp < iterations; intTmp++)
			{

				iterator[0] = Convert.ToByte((intTmp & 0x000000FF) >> 0);
				iterator[1] = Convert.ToByte((intTmp & 0x0000FF00) >> 8);
				iterator[2] = Convert.ToByte((intTmp & 0x00FF0000) >> 16);
				iterator[3] = Convert.ToByte((intTmp & 0xFF000000) >> 24);

				generatedKey = concatByteArrays(iterator, generatedKey);
				generatedKey = sha1.ComputeHash(generatedKey);
			}

			documentProtection.Add(new XAttribute(XName.Get("cryptProviderType", DocX.w.NamespaceName), "rsaFull"));
			documentProtection.Add(new XAttribute(XName.Get("cryptAlgorithmClass", DocX.w.NamespaceName), "hash"));
			documentProtection.Add(new XAttribute(XName.Get("cryptAlgorithmType", DocX.w.NamespaceName), "typeAny"));
			documentProtection.Add(new XAttribute(XName.Get("cryptAlgorithmSid", DocX.w.NamespaceName), "4"));          // SHA1
			documentProtection.Add(new XAttribute(XName.Get("cryptSpinCount", DocX.w.NamespaceName), iterations.ToString()));
			documentProtection.Add(new XAttribute(XName.Get("hash", DocX.w.NamespaceName), Convert.ToBase64String(generatedKey)));
			documentProtection.Add(new XAttribute(XName.Get("salt", DocX.w.NamespaceName), Convert.ToBase64String(arrSalt)));

			settings.Root.AddFirst(documentProtection);
		}

		private byte[] concatByteArrays(byte[] array1, byte[] array2)
		{
			byte[] result = new byte[array1.Length + array2.Length];
			Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
			Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
			return result;
		}

		/// <summary>
		/// Remove editing protection from this document.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Remove any editing restrictions that are imposed on this document.
		///     document.RemoveProtection();
		///
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AddProtection(EditRestrictions)"/>
		/// <seealso cref="GetProtectionType"/>
		/// <seealso cref="isProtected"/>
		public void RemoveProtection()
		{
			// Remove every node of type documentProtection.
			settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).Remove();
		}

		/// <summary></summary>
		public PageLayout PageLayout
		{
			get
			{
				XElement sectPr = Xml.Element(XName.Get("sectPr", DocX.w.NamespaceName));
				if (sectPr == null)
				{
					Xml.SetElementValue(XName.Get("sectPr", DocX.w.NamespaceName), string.Empty);
					sectPr = Xml.Element(XName.Get("sectPr", DocX.w.NamespaceName));
				}
				return new PageLayout(this, sectPr);
			}
		}

		/// <summary>
		/// Returns a collection of Headers in this Document.
		/// A document typically contains three Headers.
		/// A default one (odd), one for the first page and one for even pages.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Add header support to this document.
		///    document.AddHeaders();
		///
		///    // Get a collection of all headers in this document.
		///    Headers headers = document.Headers;
		///
		///    // The header used for the first page of this document.
		///    Header first = headers.first;
		///
		///    // The header used for odd pages of this document.
		///    Header odd = headers.odd;
		///
		///    // The header used for even pages of this document.
		///    Header even = headers.even;
		/// }
		/// </code>
		/// </example>
		public Headers Headers
		{
			get
			{
				return headers;
			}
		}
		private Headers headers;

		/// <summary>
		/// Returns a collection of Footers in this Document.
		/// A document typically contains three Footers.
		/// A default one (odd), one for the first page and one for even pages.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Add footer support to this document.
		///    document.AddFooters();
		///
		///    // Get a collection of all footers in this document.
		///    Footers footers = document.Footers;
		///
		///    // The footer used for the first page of this document.
		///    Footer first = footers.first;
		///
		///    // The footer used for odd pages of this document.
		///    Footer odd = footers.odd;
		///
		///    // The footer used for even pages of this document.
		///    Footer even = footers.even;
		/// }
		/// </code>
		/// </example>
		public Footers Footers
		{
			get
			{
				return footers;
			}
		}

		private Footers footers;

		/// <summary>
		/// Should the Document use different Headers and Footers for odd and even pages?
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add header support to this document.
		///     document.AddHeaders();
		///
		///     // Get a collection of all headers in this document.
		///     Headers headers = document.Headers;
		///
		///     // The header used for odd pages of this document.
		///     Header odd = headers.odd;
		///
		///     // The header used for even pages of this document.
		///     Header even = headers.even;
		///
		///     // Force the document to use a different header for odd and even pages.
		///     document.DifferentOddAndEvenPages = true;
		///
		///     // Content can be added to the Headers in the same manor that it would be added to the main document.
		///     Paragraph p1 = odd.InsertParagraph();
		///     p1.Append("This is the odd pages header.");
		///     
		///     Paragraph p2 = even.InsertParagraph();
		///     p2.Append("This is the even pages header.");
		///
		///     // Save all changes to this document.
		///     document.Save();    
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public bool DifferentOddAndEvenPages
		{
			get
			{
				XDocument settings;
				using (TextReader tr = new StreamReader(settingsPart.GetStream()))
					settings = XDocument.Load(tr);

				XElement evenAndOddHeaders = settings.Root.Element(w + "evenAndOddHeaders");

				return evenAndOddHeaders != null;
			}

			set
			{
				XDocument settings;
				using (TextReader tr = new StreamReader(settingsPart.GetStream()))
					settings = XDocument.Load(tr);

				XElement evenAndOddHeaders = settings.Root.Element(w + "evenAndOddHeaders");
				if (evenAndOddHeaders == null)
				{
					if (value)
						settings.Root.AddFirst(new XElement(w + "evenAndOddHeaders"));
				}
				else
				{
					if (!value)
						evenAndOddHeaders.Remove();
				}

				using (TextWriter tw = new StreamWriter(new PackagePartStream(settingsPart.GetStream())))
					settings.Save(tw);
			}
		}

		/// <summary>
		/// Should the Document use an independent Header and Footer for the first page?
		/// </summary>
		/// <example>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add header support to this document.
		///     document.AddHeaders();
		///
		///     // The header used for the first page of this document.
		///     Header first = document.Headers.first;
		///
		///     // Force the document to use a different header for first page.
		///     document.DifferentFirstPage = true;
		///     
		///     // Content can be added to the Headers in the same manor that it would be added to the main document.
		///     Paragraph p = first.InsertParagraph();
		///     p.Append("This is the first pages header.");
		///
		///     // Save all changes to this document.
		///     document.Save();    
		/// }// Release this document from memory.
		/// </example>
		public bool DifferentFirstPage
		{
			get
			{
				XElement body = mainDoc.Root.Element(w + "body");
				XElement sectPr = body.Element(w + "sectPr");

				if (sectPr != null)
				{
					XElement titlePg = sectPr.Element(w + "titlePg");
					if (titlePg != null)
						return true;
				}

				return false;
			}

			set
			{
				XElement body = mainDoc.Root.Element(w + "body");
				XElement sectPr = null;
				XElement titlePg = null;

				if (sectPr == null)
					body.Add(new XElement(w + "sectPr", string.Empty));

				sectPr = body.Element(w + "sectPr");

				titlePg = sectPr.Element(w + "titlePg");
				if (titlePg == null)
				{
					if (value)
						sectPr.Add(new XElement(w + "titlePg", string.Empty));
				}
				else
				{
					if (!value)
						titlePg.Remove();
				}
			}
		}

		private Header GetHeaderByType(string type)
		{
			return (Header)GetHeaderOrFooterByType(type, true);
		}

		private Footer GetFooterByType(string type)
		{
			return (Footer)GetHeaderOrFooterByType(type, false);
		}

		private object GetHeaderOrFooterByType(string type, bool isHeader)
		{
			// Switch which handles either case Header\Footer, this just cuts down on code duplication.
			string reference = "footerReference";
			if (isHeader)
				reference = "headerReference";

			// Get the Id of the [default, even or first] [Header or Footer]

			string Id =
			(
				from e in mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
				where (e.Name.LocalName == reference) && (e.Attribute(w + "type").Value == type)
				select e.Attribute(r + "id").Value
			).LastOrDefault();

			if (Id != null)
			{
				// Get the Xml file for this Header or Footer.
				Uri partUri = mainPart.GetRelationship(Id).TargetUri;

				// Weird problem with PackaePart API.
				if (!partUri.OriginalString.StartsWith("/word/"))
					partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);

				// Get the Part and open a stream to get the Xml file.
				PackagePart part = package.GetPart(partUri);

				XDocument doc;
				using (TextReader tr = new StreamReader(part.GetStream()))
				{
					doc = XDocument.Load(tr);

					// Header and Footer extend Container.
					Container c;
					if (isHeader)
						c = new Header(this, doc.Element(w + "hdr"), part);
					else
						c = new Footer(this, doc.Element(w + "ftr"), part);

					return c;
				}
			}

			// If we got this far something went wrong.
			return null;
		}

		/// <summary></summary>
		/// <returns></returns>
		public List<Section> GetSections()
		{
			var allParas = Paragraphs;
			var parasInASection = new List<Paragraph>();
			var sections = new List<Section>();
			foreach (var para in allParas)
			{
				var sectionInPara = para.Xml.Descendants().FirstOrDefault(s => s.Name.LocalName == "sectPr");
				if (sectionInPara == null)
				{
					parasInASection.Add(para);
				}
				else
				{
					parasInASection.Add(para);
					var section = new Section(Document, sectionInPara) { SectionParagraphs = parasInASection };
					sections.Add(section);
					parasInASection = new List<Paragraph>();
				}
			}
			XElement body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
			XElement baseSectionXml = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
			var baseSection = new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection };
			sections.Add(baseSection);
			return sections;
		}

		// Get the word\settings.xml part
		internal PackagePart settingsPart;

		internal PackagePart endnotesPart;

		internal PackagePart footnotesPart;

		internal PackagePart stylesPart;

		internal PackagePart stylesWithEffectsPart;

		internal PackagePart numberingPart;

		internal PackagePart fontTablePart;

		#region Internal variables defined foreach DocX object
		// Object representation of the .docx
		internal Package package;

		// The mainDocument is loaded into a XDocument object for easy querying and editing
		internal XDocument mainDoc;
		internal XDocument settings;
		internal XDocument endnotes;
		internal XDocument footnotes;
		internal XDocument styles;
		internal XDocument stylesWithEffects;
		internal XDocument numbering;
		internal XDocument fontTable;
		internal XDocument header1;
		internal XDocument header2;
		internal XDocument header3;

		// A lookup for the Paragraphs in this document.
		internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();
		// Every document is stored in a MemoryStream, all edits made to a document are done in memory.
		internal MemoryStream memoryStream;
		// The filename that this document was loaded from
		internal string filename;
		// The stream that this document was loaded from
		internal Stream stream;
		#endregion

		internal DocX(DocX document, XElement xml) : base(document, xml)
		{
		}

		/// <summary>
		/// Returns a list of Images in this document.
		/// </summary>
		/// <example>
		/// Get the unique Id of every Image in this document.
		/// <code>
		/// // Load a document.
		/// DocX document = DocX.Load(@"C:\Example\Test.docx");
		///
		/// // Loop through each Image in this document.
		/// foreach (Novacode.Image i in document.Images)
		/// {
		///     // Get the unique Id which identifies this Image.
		///     string uniqueId = i.Id;
		/// }
		///
		/// </code>
		/// </example>
		/// <seealso cref="AddImage(string)"/>
		/// <seealso cref="AddImage(Stream)"/>
		/// <seealso cref="Paragraph.Pictures"/>
		/// <seealso cref="Paragraph.InsertPicture"/>
		public List<Image> Images
		{
			get
			{
				PackageRelationshipCollection imageRelationships = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				if (imageRelationships.Count() > 0)
				{
					return
					(
						from i in imageRelationships
						select new Image(this, i)
					).ToList();
				}

				return new List<Image>();
			}
		}

		/// <summary>
		/// Returns a list of custom properties in this document.
		/// </summary>
		/// <example>
		/// Method 1: Get the name, type and value of each CustomProperty in this document.
		/// <code>
		/// // Load Example.docx
		/// DocX document = DocX.Load(@"C:\Example\Test.docx");
		///
		/// /*
		///  * No two custom properties can have the same name,
		///  * so a Dictionary is the perfect data structure to store them in.
		///  * Each custom property can be accessed using its name.
		///  */
		/// foreach (string name in document.CustomProperties.Keys)
		/// {
		///     // Grab a custom property using its name.
		///     CustomProperty cp = document.CustomProperties[name];
		///
		///     // Write this custom properties details to Console.
		///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
		/// }
		///
		/// Console.WriteLine("Press any key...");
		///
		/// // Wait for the user to press a key before closing the Console.
		/// Console.ReadKey();
		/// </code>
		/// </example>
		/// <example>
		/// Method 2: Get the name, type and value of each CustomProperty in this document.
		/// <code>
		/// // Load Example.docx
		/// DocX document = DocX.Load(@"C:\Example\Test.docx");
		/// 
		/// /*
		///  * No two custom properties can have the same name,
		///  * so a Dictionary is the perfect data structure to store them in.
		///  * The values of this Dictionary are CustomProperties.
		///  */
		/// foreach (CustomProperty cp in document.CustomProperties.Values)
		/// {
		///     // Write this custom properties details to Console.
		///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
		/// }
		///
		/// Console.WriteLine("Press any key...");
		///
		/// // Wait for the user to press a key before closing the Console.
		/// Console.ReadKey();
		/// </code>
		/// </example>
		/// <seealso cref="AddCustomProperty"/>
		public Dictionary<string, CustomProperty> CustomProperties
		{
			get
			{
				if (package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
				{
					PackagePart docProps_custom = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
					XDocument customPropDoc;
					using (TextReader tr = new StreamReader(docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
						customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

					// Get all of the custom properties in this document
					return
					(
						from p in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
						let Name = p.Attribute(XName.Get("name")).Value
						let Type = p.Descendants().Single().Name.LocalName
						let Value = p.Descendants().Single().Value
						select new CustomProperty(Name, Type, Value)
					).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
				}

				return new Dictionary<string, CustomProperty>();
			}
		}

		///<summary>Returns the list of document core properties with corresponding values</summary>
		public Dictionary<string, string> CoreProperties
		{
			get
			{
				if (package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
				{
					PackagePart docProps_Core = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
					XDocument corePropDoc;
					using (TextReader tr = new StreamReader(docProps_Core.GetStream(FileMode.Open, FileAccess.Read))) corePropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
					// Get all of the core properties in this document
					return (from docProperty in corePropDoc.Root.Elements()
							select
							  new KeyValuePair<string, string>(
							  string.Format(
								"{0}:{1}",
								corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace),
								docProperty.Name.LocalName),
							  docProperty.Value)).ToDictionary(p => p.Key, v => v.Value);
				}
				return new Dictionary<string, string>();
			}
		}

		/// <summary>
		/// Get the Text of this document.
		/// </summary>
		/// <example>
		/// Write to Console the Text from this document.
		/// <code>
		/// // Load a document
		/// DocX document = DocX.Load(@"C:\Example\Test.docx");
		///
		/// // Get the text of this document.
		/// string text = document.Text;
		///
		/// // Write the text of this document to Console.
		/// Console.Write(text);
		///
		/// // Wait for the user to press a key before closing the console window.
		/// Console.ReadKey();
		/// </code>
		/// </example>
		public string Text
		{
			get
			{
				return HelperFunctions.GetText(Xml);
			}
		}
		/// <summary>
		/// Get the text of each footnote from this document
		/// </summary>
		public IEnumerable<string> FootnotesText
		{
			get
			{
				foreach (XElement footnote in footnotes.Root.Elements(w + "footnote"))
				{
					yield return HelperFunctions.GetText(footnote);
				}
			}
		}

		/// <summary>
		/// Get the text of each endnote from this document
		/// </summary>
		public IEnumerable<string> EndnotesText
		{
			get
			{
				foreach (XElement endnote in endnotes.Root.Elements(w + "endnote"))
				{
					yield return HelperFunctions.GetText(endnote);
				}
			}
		}



		internal string GetCollectiveText(List<PackagePart> list)
		{
			string text = string.Empty;

			foreach (var hp in list)
			{
				using (TextReader tr = new StreamReader(hp.GetStream()))
				{
					XDocument d = XDocument.Load(tr);

					StringBuilder sb = new StringBuilder();

					// Loop through each text item in this run
					foreach (XElement descendant in d.Descendants())
					{
						switch (descendant.Name.LocalName)
						{
							case "tab":
								sb.Append("\t");
								break;
							case "br":
								sb.Append("\n");
								break;
							case "t":
								goto case "delText";
							case "delText":
								sb.Append(descendant.Value);
								break;
							default: break;
						}
					}

					text += "\n" + sb.ToString();
				}
			}

			return text;
		}

		/// <summary>
		/// Insert the contents of another document at the end of this document. 
		/// </summary>
		/// <param name="remote_document">The document to insert at the end of this document.</param>
		/// <param name="append">If true, document is inserted at the end, otherwise document is inserted at the beginning.</param>
		/// <example>
		/// Create a new document and insert an old document into it.
		/// <code>
		/// // Create a new document.
		/// using (DocX newDocument = DocX.Create(@"NewDocument.docx"))
		/// {
		///     // Load an old document.
		///     using (DocX oldDocument = DocX.Load(@"OldDocument.docx"))
		///     {
		///         // Insert the old document into the new document.
		///         newDocument.InsertDocument(oldDocument);
		///
		///         // Save the new document.
		///         newDocument.Save();
		///     }// Release the old document from memory.
		/// }// Release the new document from memory.
		/// </code>
		/// <remarks>
		/// If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document. In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
		/// </remarks>
		/// </example>
		public void InsertDocument(DocX remote_document, bool append = true)
		{
			// We don't want to effect the origional XDocument, so create a new one from the old one.
			XDocument remote_mainDoc = new XDocument(remote_document.mainDoc);

			XDocument remote_footnotes = null;
			if (remote_document.footnotes != null)
				remote_footnotes = new XDocument(remote_document.footnotes);

			XDocument remote_endnotes = null;
			if (remote_document.endnotes != null)
				remote_endnotes = new XDocument(remote_document.endnotes);

			// Remove all header and footer references.
			remote_mainDoc.Descendants(XName.Get("headerReference", DocX.w.NamespaceName)).Remove();
			remote_mainDoc.Descendants(XName.Get("footerReference", DocX.w.NamespaceName)).Remove();

			// Get the body of the remote document.
			XElement remote_body = remote_mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));

			// Every file that is missing from the local document will have to be copied, every file that already exists will have to be merged.
			PackagePartCollection ppc = remote_document.package.GetParts();

			List<String> ignoreContentTypes = new List<string>
			{
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
				"application/vnd.openxmlformats-package.core-properties+xml",
				"application/vnd.openxmlformats-officedocument.extended-properties+xml",
				"application/vnd.openxmlformats-package.relationships+xml",
			};

			List<String> imageContentTypes = new List<string>
			{
				"image/jpeg",
				"image/jpg",
				"image/png",
				"image/bmp",
				"image/gif",
				"image/tiff",
				"image/icon",
				"image/pcx",
				"image/emf",
				"image/wmf"
			};
			// Check if each PackagePart pp exists in this document.
			foreach (PackagePart remote_pp in ppc)
			{
				if (ignoreContentTypes.Contains(remote_pp.ContentType) || imageContentTypes.Contains(remote_pp.ContentType))
					continue;

				// If this external PackagePart already exits then we must merge them.
				if (package.PartExists(remote_pp.Uri))
				{
					PackagePart local_pp = package.GetPart(remote_pp.Uri);
					switch (remote_pp.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							merge_customs(remote_pp, local_pp, remote_mainDoc);
							break;

						// Merge footnotes (and endnotes) before merging styles, then set the remote_footnotes to the just updated footnotes
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							merge_footnotes(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes);
							remote_footnotes = footnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							merge_endnotes(remote_pp, local_pp, remote_mainDoc, remote_document, remote_endnotes);
							remote_endnotes = endnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							merge_styles(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes);
							break;

						// Merge styles after merging the footnotes, so the changes will be applied to the correct document/footnotes
						case "application/vnd.ms-word.stylesWithEffects+xml":
							merge_styles(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							merge_fonts(remote_pp, local_pp, remote_mainDoc, remote_document);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							merge_numbering(remote_pp, local_pp, remote_mainDoc, remote_document);
							break;

						default:
							break;
					}
				}

				// If this external PackagePart does not exits in the internal document then we can simply copy it.
				else
				{
					var packagePart = clonePackagePart(remote_pp);
					switch (remote_pp.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							endnotesPart = packagePart;
							endnotes = remote_endnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							footnotesPart = packagePart;
							footnotes = remote_footnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							stylesPart = packagePart;
							using (TextReader tr = new StreamReader(stylesPart.GetStream()))
								styles = XDocument.Load(tr);
							break;

						case "application/vnd.ms-word.stylesWithEffects+xml":
							stylesWithEffectsPart = packagePart;
							using (TextReader tr = new StreamReader(stylesWithEffectsPart.GetStream()))
								stylesWithEffects = XDocument.Load(tr);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							fontTablePart = packagePart;
							using (TextReader tr = new StreamReader(fontTablePart.GetStream()))
								fontTable = XDocument.Load(tr);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							numberingPart = packagePart;
							using (TextReader tr = new StreamReader(numberingPart.GetStream()))
								numbering = XDocument.Load(tr);
							break;

					}

					clonePackageRelationship(remote_document, remote_pp, remote_mainDoc);
				}
			}

			foreach (var hyperlink_rel in remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
			{
				var old_rel_Id = hyperlink_rel.Id;
				var new_rel_Id = mainPart.CreateRelationship(hyperlink_rel.TargetUri, hyperlink_rel.TargetMode, hyperlink_rel.RelationshipType).Id;
				var hyperlink_refs = remote_mainDoc.Descendants(XName.Get("hyperlink", DocX.w.NamespaceName));
				foreach (var hyperlink_ref in hyperlink_refs)
				{
					XAttribute a0 = hyperlink_ref.Attribute(XName.Get("id", DocX.r.NamespaceName));
					if (a0 != null && a0.Value == old_rel_Id)
					{
						a0.SetValue(new_rel_Id);
					}
				}
			}

			////ole object links
			foreach (var oleObject_rel in remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"))
			{
				var old_rel_Id = oleObject_rel.Id;
				var new_rel_Id = mainPart.CreateRelationship(oleObject_rel.TargetUri, oleObject_rel.TargetMode, oleObject_rel.RelationshipType).Id;
				var oleObject_refs = remote_mainDoc.Descendants(XName.Get("OLEObject", "urn:schemas-microsoft-com:office:office"));
				foreach (var oleObject_ref in oleObject_refs)
				{
					XAttribute a0 = oleObject_ref.Attribute(XName.Get("id", DocX.r.NamespaceName));
					if (a0 != null && a0.Value == old_rel_Id)
					{
						a0.SetValue(new_rel_Id);
					}
				}
			}


			foreach (PackagePart remote_pp in ppc)
			{
				if (imageContentTypes.Contains(remote_pp.ContentType))
				{
					merge_images(remote_pp, remote_document, remote_mainDoc, remote_pp.ContentType);
				}
			}

			int id = 0;
			var local_docPrs = mainDoc.Root.Descendants(XName.Get("docPr", DocX.wp.NamespaceName));
			foreach (var local_docPr in local_docPrs)
			{
				XAttribute a_id = local_docPr.Attribute(XName.Get("id"));
				int a_id_value;
				if (a_id != null && int.TryParse(a_id.Value, out a_id_value))
					if (a_id_value > id)
						id = a_id_value;
			}
			id++;

			// docPr must be sequential
			var docPrs = remote_body.Descendants(XName.Get("docPr", DocX.wp.NamespaceName));
			foreach (var docPr in docPrs)
			{
				docPr.SetAttributeValue(XName.Get("id"), id);
				id++;
			}

			// Add the remote documents contents to this document.
			XElement local_body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));
			if (append)
				local_body.Add(remote_body.Elements());
			else
				local_body.AddFirst(remote_body.Elements());

			// Copy any missing root attributes to the local document.
			foreach (XAttribute a in remote_mainDoc.Root.Attributes())
			{
				if (mainDoc.Root.Attribute(a.Name) == null)
				{
					mainDoc.Root.SetAttributeValue(a.Name, a.Value);
				}
			}

		}

		private void merge_images(PackagePart remote_pp, DocX remote_document, XDocument remote_mainDoc, String contentType)
		{
			// Before doing any other work, check to see if this image is actually referenced in the document.
			// In my testing I have found cases of Images inside documents that are not referenced
			var remote_rel = remote_document.mainPart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault();
			if (remote_rel == null)
			{
				remote_rel = remote_document.mainPart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString)).FirstOrDefault();
				if (remote_rel == null)
					return;
			}
			String remote_Id = remote_rel.Id;

			String remote_hash = ComputeMD5HashString(remote_pp.GetStream());
			var image_parts = package.GetParts().Where(pp => pp.ContentType.Equals(contentType));

			bool found = false;
			foreach (var part in image_parts)
			{
				String local_hash = ComputeMD5HashString(part.GetStream());
				if (local_hash.Equals(remote_hash))
				{
					// This image already exists in this document.
					found = true;

					var local_rel = mainPart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault();
					if (local_rel == null)
					{
						local_rel = mainPart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString)).FirstOrDefault();
					}
					if (local_rel != null)
					{
						String new_Id = local_rel.Id;

						// Replace all instances of remote_Id in the local document with local_Id
						var elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName));
						foreach (var elem in elems)
						{
							XAttribute embed = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName));
							if (embed != null && embed.Value == remote_Id)
							{
								embed.SetValue(new_Id);
							}
						}

						// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
						var v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName));
						foreach (var elem in v_elems)
						{
							XAttribute id = elem.Attribute(XName.Get("id", DocX.r.NamespaceName));
							if (id != null && id.Value == remote_Id)
							{
								id.SetValue(new_Id);
							}
						}
					}

					break;
				}
			}

			// This image does not exist in this document.
			if (!found)
			{
				String new_uri = remote_pp.Uri.OriginalString;
				new_uri = new_uri.Remove(new_uri.LastIndexOf("/"));
				//new_uri = new_uri.Replace("word/", "");
				new_uri += "/" + Guid.NewGuid().ToString() + contentType.Replace("image/", ".");
				if (!new_uri.StartsWith("/"))
					new_uri = "/" + new_uri;

				PackagePart new_pp = package.CreatePart(new Uri(new_uri, UriKind.Relative), remote_pp.ContentType, CompressionOption.Normal);

				using (Stream s_read = remote_pp.GetStream())
				{
					using (Stream s_write = new PackagePartStream(new_pp.GetStream(FileMode.Create)))
					{
						byte[] buffer = new byte[32768];
						int read;
						while ((read = s_read.Read(buffer, 0, buffer.Length)) > 0)
						{
							s_write.Write(buffer, 0, read);
						}
					}
				}

				PackageRelationship pr = mainPart.CreateRelationship(new Uri(new_uri, UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

				String new_Id = pr.Id;

				//Check if the remote relationship id is a default rId from Word
				Match defRelId = Regex.Match(remote_Id, @"rId\d+", RegexOptions.IgnoreCase);

				// Replace all instances of remote_Id in the local document with local_Id
				var elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName));
				foreach (var elem in elems)
				{
					XAttribute embed = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName));
					if (embed != null && embed.Value == remote_Id)
					{
						embed.SetValue(new_Id);
					}
				}

				if (!defRelId.Success)
				{
					// Replace all instances of remote_Id in the local document with local_Id
					var elems_local = mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName));
					foreach (var elem in elems_local)
					{
						XAttribute embed = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName));
						if (embed != null && embed.Value == remote_Id)
						{
							embed.SetValue(new_Id);
						}
					}


					// Replace all instances of remote_Id in the local document with local_Id
					var v_elems_local = mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName));
					foreach (var elem in v_elems_local)
					{
						XAttribute id = elem.Attribute(XName.Get("id", DocX.r.NamespaceName));
						if (id != null && id.Value == remote_Id)
						{
							id.SetValue(new_Id);
						}
					}
				}


				// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
				var v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName));
				foreach (var elem in v_elems)
				{
					XAttribute id = elem.Attribute(XName.Get("id", DocX.r.NamespaceName));
					if (id != null && id.Value == remote_Id)
					{
						id.SetValue(new_Id);
					}
				}
			}
		}

		private string ComputeMD5HashString(Stream stream)
		{
			MD5 md5 = MD5.Create();
			byte[] hash = md5.ComputeHash(stream);
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < hash.Length; i++)
				sb.Append(hash[i].ToString("X2"));
			return sb.ToString();
		}

		private void merge_endnotes(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_endnotes)
		{
			IEnumerable<int> ids =
			(
				from d in endnotes.Root.Descendants()
				where d.Name.LocalName == "endnote"
				select int.Parse(d.Attribute(XName.Get("id", DocX.w.NamespaceName)).Value)
			);

			int max_id = ids.Max() + 1;
			var endnoteReferences = remote_mainDoc.Descendants(XName.Get("endnoteReference", DocX.w.NamespaceName));

			foreach (var endnote in remote_endnotes.Root.Elements().OrderBy(fr => fr.Attribute(XName.Get("id", DocX.r.NamespaceName))).Reverse())
			{
				XAttribute id = endnote.Attribute(XName.Get("id", DocX.w.NamespaceName));
				int i;
				if (id != null && int.TryParse(id.Value, out i))
				{
					if (i > 0)
					{
						foreach (var endnoteRef in endnoteReferences)
						{
							XAttribute a = endnoteRef.Attribute(XName.Get("id", DocX.w.NamespaceName));
							if (a != null && int.Parse(a.Value).Equals(i))
							{
								a.SetValue(max_id);
							}
						}

						// We care about copying this footnote.
						endnote.SetAttributeValue(XName.Get("id", DocX.w.NamespaceName), max_id);
						endnotes.Root.Add(endnote);
						max_id++;
					}
				}
			}
		}

		private void merge_footnotes(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes)
		{
			IEnumerable<int> ids =
			(
				from d in footnotes.Root.Descendants()
				where d.Name.LocalName == "footnote"
				select int.Parse(d.Attribute(XName.Get("id", DocX.w.NamespaceName)).Value)
			);

			int max_id = ids.Max() + 1;
			var footnoteReferences = remote_mainDoc.Descendants(XName.Get("footnoteReference", DocX.w.NamespaceName));

			foreach (var footnote in remote_footnotes.Root.Elements().OrderBy(fr => fr.Attribute(XName.Get("id", DocX.r.NamespaceName))).Reverse())
			{
				XAttribute id = footnote.Attribute(XName.Get("id", DocX.w.NamespaceName));
				int i;
				if (id != null && int.TryParse(id.Value, out i))
				{
					if (i > 0)
					{
						foreach (var footnoteRef in footnoteReferences)
						{
							XAttribute a = footnoteRef.Attribute(XName.Get("id", DocX.w.NamespaceName));
							if (a != null && int.Parse(a.Value).Equals(i))
							{
								a.SetValue(max_id);
							}
						}

						// We care about copying this footnote.
						footnote.SetAttributeValue(XName.Get("id", DocX.w.NamespaceName), max_id);
						footnotes.Root.Add(footnote);
						max_id++;
					}
				}
			}
		}

		private void merge_customs(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc)
		{
			// Get the remote documents custom.xml file.
			XDocument remote_custom_document;
			using (TextReader tr = new StreamReader(remote_pp.GetStream()))
				remote_custom_document = XDocument.Load(tr);

			// Get the local documents custom.xml file.
			XDocument local_custom_document;
			using (TextReader tr = new StreamReader(local_pp.GetStream()))
				local_custom_document = XDocument.Load(tr);

			IEnumerable<int> pids =
			(
				from d in remote_custom_document.Root.Descendants()
				where d.Name.LocalName == "property"
				select int.Parse(d.Attribute(XName.Get("pid")).Value)
			);

			int pid = pids.Max() + 1;

			foreach (XElement remote_property in remote_custom_document.Root.Elements())
			{
				bool found = false;
				foreach (XElement local_property in local_custom_document.Root.Elements())
				{
					XAttribute remote_property_name = remote_property.Attribute(XName.Get("name"));
					XAttribute local_property_name = local_property.Attribute(XName.Get("name"));

					if (remote_property != null && local_property_name != null && remote_property_name.Value.Equals(local_property_name.Value))
						found = true;
				}

				if (!found)
				{
					remote_property.SetAttributeValue(XName.Get("pid"), pid);
					local_custom_document.Root.Add(remote_property);

					pid++;
				}
			}

			// Save the modified local custom styles.xml file.
			using (TextWriter tw = new StreamWriter(new PackagePartStream(local_pp.GetStream(FileMode.Create, FileAccess.Write))))
				local_custom_document.Save(tw, SaveOptions.None);
		}

		private void merge_numbering(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote)
		{
			// Add each remote numbering to this document.
			IEnumerable<XElement> remote_abstractNums = remote.numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName));
			int guidd = 0;
			foreach (var an in remote_abstractNums)
			{
				XAttribute a = an.Attribute(XName.Get("abstractNumId", DocX.w.NamespaceName));
				if (a != null)
				{
					int i;
					if (int.TryParse(a.Value, out i))
					{
						if (i > guidd)
							guidd = i;
					}
				}
			}
			guidd++;

			IEnumerable<XElement> remote_nums = remote.numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName));
			int guidd2 = 0;
			foreach (var an in remote_nums)
			{
				XAttribute a = an.Attribute(XName.Get("numId", DocX.w.NamespaceName));
				if (a != null)
				{
					int i;
					if (int.TryParse(a.Value, out i))
					{
						if (i > guidd2)
							guidd2 = i;
					}
				}
			}
			guidd2++;

			foreach (XElement remote_abstractNum in remote_abstractNums)
			{
				XAttribute abstractNumId = remote_abstractNum.Attribute(XName.Get("abstractNumId", DocX.w.NamespaceName));
				if (abstractNumId != null)
				{
					String abstractNumIdValue = abstractNumId.Value;
					abstractNumId.SetValue(guidd);

					foreach (XElement remote_num in remote_nums)
					{
						var numIds = remote_mainDoc.Descendants(XName.Get("numId", DocX.w.NamespaceName));
						foreach (var numId in numIds)
						{
							XAttribute attr = numId.Attribute(XName.Get("val", DocX.w.NamespaceName));
							if (attr != null && attr.Value.Equals(remote_num.Attribute(XName.Get("numId", DocX.w.NamespaceName)).Value))
							{
								attr.SetValue(guidd2);
							}

						}
						remote_num.SetAttributeValue(XName.Get("numId", DocX.w.NamespaceName), guidd2);

						XElement e = remote_num.Element(XName.Get("abstractNumId", DocX.w.NamespaceName));
						if (e != null)
						{
							XAttribute a2 = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
							if (a2 != null && a2.Value.Equals(abstractNumIdValue))
								a2.SetValue(guidd);
						}

						guidd2++;
					}
				}

				guidd++;
			}

			// Checking whether there were more than 0 elements, helped me get rid of exceptions thrown while using InsertDocument
			if (numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName)).Count() > 0)
				numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName)).Last().AddAfterSelf(remote_abstractNums);

			if (numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName)).Count() > 0)
				numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName)).Last().AddAfterSelf(remote_nums);
		}

		private void merge_fonts(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote)
		{
			// Add each remote font to this document.
			IEnumerable<XElement> remote_fonts = remote.fontTable.Root.Elements(XName.Get("font", DocX.w.NamespaceName));
			IEnumerable<XElement> local_fonts = fontTable.Root.Elements(XName.Get("font", DocX.w.NamespaceName));

			foreach (XElement remote_font in remote_fonts)
			{
				bool flag_addFont = true;
				foreach (XElement local_font in local_fonts)
				{
					if (local_font.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value == remote_font.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value)
					{
						flag_addFont = false;
						break;
					}
				}

				if (flag_addFont)
				{
					fontTable.Root.Add(remote_font);
				}
			}
		}

		private void merge_styles(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes, XDocument remote_endnotes)
		{
			Dictionary<String, String> local_styles = new Dictionary<string, string>();
			foreach (XElement local_style in styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName)))
			{
				XElement temp = new XElement(local_style);
				XAttribute styleId = temp.Attribute(XName.Get("styleId", DocX.w.NamespaceName));
				String value = styleId.Value;
				styleId.Remove();
				String key = Regex.Replace(temp.ToString(), @"\s+", "");
				if (!local_styles.ContainsKey(key)) local_styles.Add(key, value);
			}

			// Add each remote style to this document.
			IEnumerable<XElement> remote_styles = remote.styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName));
			foreach (XElement remote_style in remote_styles)
			{
				XElement temp = new XElement(remote_style);
				XAttribute styleId = temp.Attribute(XName.Get("styleId", DocX.w.NamespaceName));
				String value = styleId.Value;
				styleId.Remove();
				String key = Regex.Replace(temp.ToString(), @"\s+", "");
				String guuid;

				// Check to see if the local document already contains the remote style.
				if (local_styles.ContainsKey(key))
				{
					String local_value;
					local_styles.TryGetValue(key, out local_value);

					// If the styleIds are the same then nothing needs to be done.
					if (local_value == value)
						continue;

					// All we need to do is update the styleId.
					else
					{
						guuid = local_value;
					}
				}
				else
				{
					guuid = Guid.NewGuid().ToString();
					// Set the styleId in the remote_style to this new Guid
					// [Fixed the issue that my document referred to a new Guid while my styles still had the old value ("Titel")]
					remote_style.SetAttributeValue(XName.Get("styleId", DocX.w.NamespaceName), guuid);
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName)))
				{
					XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
					{
						e_styleId.SetValue(guuid);
					}
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName)))
				{
					XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
					{
						e_styleId.SetValue(guuid);
					}
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(XName.Get("tblStyle", DocX.w.NamespaceName)))
				{
					XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
					{
						e_styleId.SetValue(guuid);
					}
				}

				if (remote_endnotes != null)
				{
					foreach (XElement e in remote_endnotes.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName)))
					{
						XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
						if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
						{
							e_styleId.SetValue(guuid);
						}
					}

					foreach (XElement e in remote_endnotes.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName)))
					{
						XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
						if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
						{
							e_styleId.SetValue(guuid);
						}
					}
				}

				if (remote_footnotes != null)
				{
					foreach (XElement e in remote_footnotes.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName)))
					{
						XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
						if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
						{
							e_styleId.SetValue(guuid);
						}
					}

					foreach (XElement e in remote_footnotes.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName)))
					{
						XAttribute e_styleId = e.Attribute(XName.Get("val", DocX.w.NamespaceName));
						if (e_styleId != null && e_styleId.Value.Equals(styleId.Value))
						{
							e_styleId.SetValue(guuid);
						}
					}
				}

				// Make sure they don't clash by using a uuid.
				styleId.SetValue(guuid);
				styles.Root.Add(remote_style);
			}
		}

		/// <summary></summary>
		/// <param name="remote_document"></param>
		/// <param name="pp"></param>
		/// <param name="remote_mainDoc"></param>
		protected void clonePackageRelationship(DocX remote_document, PackagePart pp, XDocument remote_mainDoc)
		{
			string url = pp.Uri.OriginalString.Replace("/", "");
			var remote_rels = remote_document.mainPart.GetRelationships();
			foreach (var remote_rel in remote_rels)
			{
				if (url.Equals("word" + remote_rel.TargetUri.OriginalString.Replace("/", "")))
				{
					String remote_Id = remote_rel.Id;
					String local_Id = mainPart.CreateRelationship(remote_rel.TargetUri, remote_rel.TargetMode, remote_rel.RelationshipType).Id;

					// Replace all instances of remote_Id in the local document with local_Id
					var elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName));
					foreach (var elem in elems)
					{
						XAttribute embed = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName));
						if (embed != null && embed.Value == remote_Id)
						{
							embed.SetValue(local_Id);
						}
					}

					// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
					var v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName));
					foreach (var elem in v_elems)
					{
						XAttribute id = elem.Attribute(XName.Get("id", DocX.r.NamespaceName));
						if (id != null && id.Value == remote_Id)
						{
							id.SetValue(local_Id);
						}
					}
					break;
				}
			}
		}

		/// <summary></summary>
		/// <param name="pp"></param>
		/// <returns></returns>
		protected PackagePart clonePackagePart(PackagePart pp)
		{
			PackagePart new_pp = package.CreatePart(pp.Uri, pp.ContentType, CompressionOption.Normal);

			using (Stream s_read = pp.GetStream())
			{
				using (Stream s_write = new PackagePartStream(new_pp.GetStream(FileMode.Create)))
				{
					byte[] buffer = new byte[32768];
					int read;
					while ((read = s_read.Read(buffer, 0, buffer.Length)) > 0)
					{
						s_write.Write(buffer, 0, read);
					}
				}
			}

			return new_pp;
		}

		/// <summary></summary>
		/// <param name="stream"></param>
		/// <returns></returns>
		protected string GetMD5HashFromStream(Stream stream)
		{
			MD5 md5 = new MD5CryptoServiceProvider();
			byte[] retVal = md5.ComputeHash(stream);
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < retVal.Length; i++) sb.Append(retVal[i].ToString("x2"));
			return sb.ToString();
		}

		/// <summary>Insert a new Table at the end of this document</summary>
		/// <param name="columnCount">The number of columns to create.</param>
		/// <param name="rowCount">The number of rows to create.</param>
		/// <returns>A new Table.</returns>
		/// <example>
		/// Insert a new Table with 2 columns and 3 rows, at the end of a document.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"C:\Example\Test.docx"))
		/// {
		///     // Create a new Table with 2 columns and 3 rows.
		///     Table newTable = document.InsertTable(2, 3);
		///
		///     // Set the design of this Table.
		///     newTable.Design = TableDesign.LightShadingAccent2;
		///
		///     // Set the column names.
		///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
		///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
		///
		///     // Fill row 1
		///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
		///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
		///
		///     // Fill row 2
		///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
		///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
		///
		///     // Save all changes made to document b.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public new Table InsertTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = base.InsertTable(rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public Table AddTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = new Table(this, HelperFunctions.CreateTable(rowCount, columnCount));
			t.mainPart = mainPart;
			return t;
		}

		/// <summary>Create a new list with a list item</summary>
		/// <param name="listText">The text of the first element in the created list.</param>
		/// <param name="level">The indentation level of the element in the list.</param>
		/// <param name="listType">The type of list to be created: Bulleted or Numbered.</param>
		/// <param name="startNumber">The number start number for the list. </param>
		/// <param name="trackChanges">Enable change tracking</param>
		/// <param name="continueNumbering">Set to true if you want to continue numbering from the previous numbered list</param>
		/// <returns>
		/// The created List. Call AddListItem(...) to add more elements to the list.
		/// Write the list to the Document with InsertList(...) once the list has all the desired 
		/// elements, otherwise the list will not be included in the working Document.
		/// </returns>
		public List AddList(string listText = null, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
		{
			return AddListItem(new List(this, null), listText, level, listType, startNumber, trackChanges, continueNumbering);
		}

		/// <summary>Add a list item to an already existing list</summary>
		/// <param name="list">The list to add the new list item to.</param>
		/// <param name="listText">The run text that should be in the new list item.</param>
		/// <param name="level">The indentation level of the new list element.</param>
		/// <param name="startNumber">The number start number for the list. </param>
		/// <param name="trackChanges">Enable change tracking</param>
		/// <param name="listType">Numbered or Bulleted list type. </param>
		/// /// <param name="continueNumbering">Set to true if you want to continue numbering from the previous numbered list</param>
		/// <returns>
		/// The created List. Call AddListItem(...) to add more elements to the list.
		/// Write the list to the Document with InsertList(...) once the list has all the desired 
		/// elements, otherwise the list will not be included in the working Document.
		/// </returns>
		public List AddListItem(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
		{
			if (startNumber.HasValue && continueNumbering) throw new InvalidOperationException("Cannot specify a start number and at the same time continue numbering from another list");
			var listToReturn = HelperFunctions.CreateItemInList(list, listText, level, listType, startNumber, trackChanges, continueNumbering);
			var lastItem = listToReturn.Items.LastOrDefault();
			if (lastItem != null) lastItem.PackagePart = mainPart;
			return listToReturn;
		}

		/// <summary>Insert list into the document</summary>
		/// <param name="list">The list to insert into the document</param>
		/// <returns>The list that was inserted into the document</returns>
		public new List InsertList(List list)
		{
			base.InsertList(list);
			return list;
		}

		/// <summary></summary>
		/// <param name="list"></param>
		/// <param name="fontFamily"></param>
		/// <param name="fontSize"></param>
		/// <returns></returns>
		public new List InsertList(List list, Font fontFamily, double fontSize)
		{
			base.InsertList(list, fontFamily, fontSize);
			return list;
		}

		/// <summary></summary>
		/// <param name="list"></param>
		/// <param name="fontSize"></param>
		/// <returns></returns>
		public new List InsertList(List list, double fontSize)
		{
			base.InsertList(list, fontSize);
			return list;
		}

		/// <summary>Insert a list at an index location in the document</summary>
		/// <param name="index">Index in document to insert the list.</param>
		/// <param name="list">The list that was inserted into the document.</param>
		/// <returns></returns>
		public new List InsertList(int index, List list)
		{
			base.InsertList(index, list);
			return list;
		}

		internal XDocument AddStylesForList()
		{
			var wordStylesUri = new Uri("/word/styles.xml", UriKind.Relative);
			// If the internal document contains no /word/styles.xml create one.
			if (!package.PartExists(wordStylesUri)) HelperFunctions.AddDefaultStylesXml(package);
			// Load the styles.xml into memory.
			XDocument wordStyles;
			using (TextReader tr = new StreamReader(package.GetPart(wordStylesUri).GetStream())) wordStyles = XDocument.Load(tr);
			bool listStyleExists = (
				from s in wordStyles.Element(w + "styles").Elements()
				let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
				where (styleId != null && styleId.Value == "ListParagraph")
				select s
			).Any();
			if (!listStyleExists)
			{
				var style = new XElement
				(
					w + "style",
					new XAttribute(w + "type", "paragraph"),
					new XAttribute(w + "styleId", "ListParagraph"),
						new XElement(w + "name", new XAttribute(w + "val", "List Paragraph")),
						new XElement(w + "basedOn", new XAttribute(w + "val", "Normal")),
						new XElement(w + "uiPriority", new XAttribute(w + "val", "34")),
						new XElement(w + "qformat"),
						new XElement(w + "rsid", new XAttribute(w + "val", "00832EE1")),
						new XElement
						(
							w + "rPr",
							new XElement(w + "ind", new XAttribute(w + "left", "720")),
							new XElement
							(
								w + "contextualSpacing"
							)
						)
				);
				wordStyles.Element(w + "styles").Add(style);
				// Save the styles document
				using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(wordStylesUri).GetStream()))) wordStyles.Save(tw);
			}
			return wordStyles;
		}

		/// <summary>Insert a Table into this document. The Table's source can be a completely different document.</summary>
		/// <param name="t">The Table to insert.</param>
		/// <param name="index">The index to insert this Table at.</param>
		/// <returns>The Table now associated with this document.</returns>
		/// <example>
		/// Extract a Table from document a and insert it into document b, at index 10.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
		/// {
		///     /* 
		///      * Insert the Table that was extracted from document a, into document b. 
		///      * This creates a new Table that is now associated with document b.
		///      */
		///     Table newTable = documentB.InsertTable(10, t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public new Table InsertTable(int index, Table t)
		{
			Table t2 = base.InsertTable(index, t);
			t2.mainPart = mainPart;
			return t2;
		}

		/// <summary>
		/// Insert a Table into this document. The Table's source can be a completely different document.
		/// </summary>
		/// <param name="t">The Table to insert.</param>
		/// <returns>The Table now associated with this document.</returns>
		/// <example>
		/// Extract a Table from document a and insert it at the end of document b.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
		/// {
		///     /* 
		///      * Insert the Table that was extracted from document a, into document b. 
		///      * This creates a new Table that is now associated with document b.
		///      */
		///     Table newTable = documentB.InsertTable(t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public new Table InsertTable(Table t)
		{
			t = base.InsertTable(t);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary>
		/// Insert a new Table at the end of this document.
		/// </summary>
		/// <param name="columnCount">The number of columns to create.</param>
		/// <param name="rowCount">The number of rows to create.</param>
		/// <param name="index">The index to insert this Table at.</param>
		/// <returns>A new Table.</returns>
		/// <example>
		/// Insert a new Table with 2 columns and 3 rows, at index 37 in this document.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Create a new Table with 3 rows and 2 columns. Insert this Table at index 37.
		///     Table newTable = document.InsertTable(37, 3, 2);
		///
		///     // Set the design of this Table.
		///     newTable.Design = TableDesign.LightShadingAccent3;
		///
		///     // Set the column names.
		///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
		///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
		///
		///     // Fill row 1
		///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
		///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
		///
		///     // Fill row 2
		///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
		///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
		///
		///     // Save all changes made to document b.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public new Table InsertTable(int index, int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");

			Table t = base.InsertTable(index, rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary>
		/// Creates a document using a Stream.
		/// </summary>
		/// <param name="stream">The Stream to create the document from.</param>
		/// <param name="documentType"></param>
		/// <returns>Returns a DocX object which represents the document.</returns>
		/// <example>
		/// Creating a document from a FileStream.
		/// <code>
		/// // Use a FileStream fs to create a new document.
		/// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
		/// {
		///     // Load the document using fs
		///     using (DocX document = DocX.Create(fs))
		///     {
		///         // Do something with the document here.
		///
		///         // Save all changes made to this document.
		///         document.Save();
		///     }// Release this document from memory.
		/// }
		/// </code>
		/// </example>
		/// <example>
		/// Creating a document in a SharePoint site.
		/// <code>
		/// using(SPSite mySite = new SPSite("http://server/sites/site"))
		/// {
		///     // Open a connection to the SharePoint site
		///     using(SPWeb myWeb = mySite.OpenWeb())
		///     {
		///         // Create a MemoryStream ms.
		///         using (MemoryStream ms = new MemoryStream())
		///         {
		///             // Create a document using ms.
		///             using (DocX document = DocX.Create(ms))
		///             {
		///                 // Do something with the document here.
		///
		///                 // Save all changes made to this document.
		///                 document.Save();
		///             }// Release this document from memory
		///
		///             // Add the document to the SharePoint site
		///             web.Files.Add("filename", ms.ToArray(), true);
		///         }
		///     }
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Load(string)"/>
		/// <seealso cref="DocX.Save()"/>
		public static DocX Create(Stream stream, DocumentTypes documentType = DocumentTypes.Document)
		{
			// Store this document in memory
			MemoryStream ms = new MemoryStream();

			// Create the docx package
			Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

			PostCreation(package, documentType);
			DocX document = DocX.Load(ms);
			document.stream = stream;
			return document;
		}

		/// <summary>
		/// Creates a document using a fully qualified or relative filename.
		/// </summary>
		/// <param name="filename">The fully qualified or relative filename.</param>
		/// <param name="documentType"></param>
		/// <returns>Returns a DocX object which represents the document.</returns>
		/// <example>
		/// <code>
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Create(@"..\Test.docx"))
		/// {
		///     // Do something with the document here.
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory
		/// </code>
		/// <code>
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Create(@"..\Test.docx"))
		/// {
		///     // Do something with the document here.
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory
		/// </code>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Load(string)"/>
		/// <seealso cref="DocX.Save()"/>
		/// </example>
		public static DocX Create(string filename, DocumentTypes documentType = DocumentTypes.Document)
		{
			// Store this document in memory
			MemoryStream ms = new MemoryStream();

			// Create the docx package
			//WordprocessingDocument wdDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
			Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

			PostCreation(package, documentType);
			DocX document = DocX.Load(ms);
			document.filename = filename;
			return document;
		}

		internal static void PostCreation(Package package, DocumentTypes documentType = DocumentTypes.Document)
		{
			XDocument mainDoc, stylesDoc, numberingDoc;

			#region MainDocumentPart
			// Create the main document part for this package
			PackagePart mainDocumentPart;
			if (documentType == DocumentTypes.Document)
			{
				mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", CompressionOption.Normal);
			}
			else
			{
				mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", CompressionOption.Normal);
			}
			package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

			// Load the document part into a XDocument object
			using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
			{
				mainDoc = XDocument.Parse
				(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                   <w:body>
                    <w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">
                        <w:pgSz w:w=""11906"" w:h=""16838""/>
                        <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>
                        <w:cols w:space=""708""/>
                        <w:docGrid w:linePitch=""360""/>
                    </w:sectPr>
                   </w:body>
                   </w:document>"
				);
			}

			// Save the main document
			using (TextWriter tw = new StreamWriter(new PackagePartStream(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))))
				mainDoc.Save(tw, SaveOptions.None);
			#endregion

			#region StylePart
			stylesDoc = HelperFunctions.AddDefaultStylesXml(package);
			#endregion

			#region NumberingPart
			numberingDoc = HelperFunctions.AddDefaultNumberingXml(package);
			#endregion

			package.Close();
		}

		internal static DocX PostLoad(ref Package package)
		{
			DocX document = new DocX(null, null);
			document.package = package;
			document.Document = document;

			#region MainDocumentPart
			document.mainPart = package.GetParts().Where
			(
					 p => p.ContentType.Equals(HelperFunctions.DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) ||
					 p.ContentType.Equals(HelperFunctions.TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)
			).Single();

			using (TextReader tr = new StreamReader(document.mainPart.GetStream(FileMode.Open, FileAccess.Read)))
				document.mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
			#endregion

			PopulateDocument(document, package);

			using (TextReader tr = new StreamReader(document.settingsPart.GetStream()))
				document.settings = XDocument.Load(tr);

			document.paragraphLookup.Clear();
			foreach (var paragraph in document.Paragraphs)
			{
				if (!document.paragraphLookup.ContainsKey(paragraph.endIndex))
					document.paragraphLookup.Add(paragraph.endIndex, paragraph);
			}

			return document;
		}

		private static void PopulateDocument(DocX document, Package package)
		{
			Headers headers = new Headers();
			headers.odd = document.GetHeaderByType("default");
			headers.even = document.GetHeaderByType("even");
			headers.first = document.GetHeaderByType("first");

			Footers footers = new Footers();
			footers.odd = document.GetFooterByType("default");
			footers.even = document.GetFooterByType("even");
			footers.first = document.GetFooterByType("first");

			//// Get the sectPr for this document.
			//XElement sect = document.mainDoc.Descendants(XName.Get("sectPr", DocX.w.NamespaceName)).Single();

			//if (sectPr != null)
			//{
			//    // Extract the even header reference
			//    var header_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
			//    string id = header_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res = document.mainPart.GetRelationship(id);
			//    string ans = res.SourceUri.OriginalString;
			//    headers.even.xml_filename = ans;

			//    // Extract the odd header reference
			//    var header_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
			//    string id2 = header_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res2 = document.mainPart.GetRelationship(id2);
			//    string ans2 = res2.SourceUri.OriginalString;
			//    headers.odd.xml_filename = ans2;

			//    // Extract the first header reference
			//    var header_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "h
			//eaderReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
			//    string id3 = header_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res3 = document.mainPart.GetRelationship(id3);
			//    string ans3 = res3.SourceUri.OriginalString;
			//    headers.first.xml_filename = ans3;

			//    // Extract the even footer reference
			//    var footer_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
			//    string id4 = footer_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res4 = document.mainPart.GetRelationship(id4);
			//    string ans4 = res4.SourceUri.OriginalString;
			//    footers.even.xml_filename = ans4;

			//    // Extract the odd footer reference
			//    var footer_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
			//    string id5 = footer_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res5 = document.mainPart.GetRelationship(id5);
			//    string ans5 = res5.SourceUri.OriginalString;
			//    footers.odd.xml_filename = ans5;

			//    // Extract the first footer reference
			//    var footer_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
			//    string id6 = footer_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			//    var res6 = document.mainPart.GetRelationship(id6);
			//    string ans6 = res6.SourceUri.OriginalString;
			//    footers.first.xml_filename = ans6;

			//}

			document.Xml = document.mainDoc.Root.Element(w + "body");
			document.headers = headers;
			document.footers = footers;
			document.settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);

			var ps = package.GetParts();

			//document.endnotesPart = HelperFunctions.GetPart();

			foreach (var rel in document.mainPart.GetRelationships())
			{
				string url = "/word/" + rel.TargetUri.OriginalString.Replace("/word/", "").Replace("file://", "");

				switch (rel.RelationshipType)
				{
					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
						document.endnotesPart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.endnotesPart.GetStream()))
							document.endnotes = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
						document.footnotesPart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.footnotesPart.GetStream()))
							document.footnotes = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
						document.stylesPart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.stylesPart.GetStream()))
							document.styles = XDocument.Load(tr);
						break;

					case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
						document.stylesWithEffectsPart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.stylesWithEffectsPart.GetStream()))
							document.stylesWithEffects = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
						document.fontTablePart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.fontTablePart.GetStream()))
							document.fontTable = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
						document.numberingPart = package.GetPart(new Uri(url, UriKind.Relative));
						using (TextReader tr = new StreamReader(document.numberingPart.GetStream()))
							document.numbering = XDocument.Load(tr);
						break;

					default:
						break;
				}
			}
		}

		/// <summary>
		/// Saves and copies the document into a new DocX object
		/// </summary>
		/// <returns>
		/// Returns a new DocX object with an identical document
		/// </returns>
		/// <example>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Save()"/>
		/// </example>
		public DocX Copy()
		{
			MemoryStream ms = new MemoryStream();
			SaveAs(ms);
			ms.Seek(0, SeekOrigin.Begin);

			return DocX.Load(ms);
		}

		/// <summary>
		/// Loads a document into a DocX object using a Stream.
		/// </summary>
		/// <param name="stream">The Stream to load the document from.</param>
		/// <returns>
		/// Returns a DocX object which represents the document.
		/// </returns>
		/// <example>
		/// Loading a document from a FileStream.
		/// <code>
		/// // Open a FileStream fs to a document.
		/// using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
		/// {
		///     // Load the document using fs.
		///     using (DocX document = DocX.Load(fs))
		///     {
		///         // Do something with the document here.
		///            
		///         // Save all changes made to the document.
		///         document.Save();
		///     }// Release this document from memory.
		/// }
		/// </code>
		/// </example>
		/// <example>
		/// Loading a document from a SharePoint site.
		/// <code>
		/// // Get the SharePoint site that you want to access.
		/// using (SPSite mySite = new SPSite("http://server/sites/site"))
		/// {
		///     // Open a connection to the SharePoint site
		///     using (SPWeb myWeb = mySite.OpenWeb())
		///     {
		///         // Grab a document stored on this site.
		///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
		///
		///         // DocX.Load requires a Stream, so open a Stream to this document.
		///         Stream str = new MemoryStream(file.OpenBinary());
		///
		///         // Load the file using the Stream str.
		///         using (DocX document = DocX.Load(str))
		///         {
		///             // Do something with the document here.
		///
		///             // Save all changes made to the document.
		///             document.Save();
		///         }// Release this document from memory.
		///     }
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="DocX.Load(string)"/>
		/// <seealso cref="DocX.Save()"/>
		public static DocX Load(Stream stream)
		{
			MemoryStream ms = new MemoryStream();

			stream.Position = 0;
			byte[] data = new byte[stream.Length];
			stream.Read(data, 0, (int)stream.Length);
			ms.Write(data, 0, (int)stream.Length);

			// Open the docx package
			Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

			DocX document = PostLoad(ref package);
			document.package = package;
			document.memoryStream = ms;
			document.stream = stream;
			return document;
		}

		/// <summary>
		/// Loads a document into a DocX object using a fully qualified or relative filename.
		/// </summary>
		/// <param name="filename">The fully qualified or relative filename.</param>
		/// <returns>
		/// Returns a DocX object which represents the document.
		/// </returns>
		/// <example>
		/// <code>
		/// // Load a document using its fully qualified filename
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Do something with the document here
		///
		///     // Save all changes made to document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// <code>
		/// // Load a document using its relative filename.
		/// using(DocX document = DocX.Load(@"..\..\Test.docx"))
		/// { 
		///     // Do something with the document here.
		///                
		///     // Save all changes made to document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Save()"/>
		/// </example>
		public static DocX Load(string filename)
		{
			if (!File.Exists(filename))
				throw new FileNotFoundException(string.Format("File could not be found {0}", filename));

			MemoryStream ms = new MemoryStream();

			using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				byte[] data = new byte[fs.Length];
				fs.Read(data, 0, (int)fs.Length);
				ms.Write(data, 0, (int)fs.Length);
			}

			// Open the docx package
			Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

			DocX document = PostLoad(ref package);
			document.package = package;
			document.filename = filename;
			document.memoryStream = ms;

			return document;
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateFilePath">The path to the document template file.</param>
		///<exception cref="FileNotFoundException">The document template file not found.</exception>
		public void ApplyTemplate(string templateFilePath)
		{
			ApplyTemplate(templateFilePath, true);
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateFilePath">The path to the document template file.</param>
		///<param name="includeContent">Whether to copy the document template text content to document.</param>
		///<exception cref="FileNotFoundException">The document template file not found.</exception>
		public void ApplyTemplate(string templateFilePath, bool includeContent)
		{
			if (!File.Exists(templateFilePath))
			{
				throw new FileNotFoundException(string.Format("File could not be found {0}", templateFilePath));
			}
			using (FileStream packageStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read))
			{
				ApplyTemplate(packageStream, includeContent);
			}
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateStream">The stream of the document template file.</param>
		public void ApplyTemplate(Stream templateStream)
		{
			ApplyTemplate(templateStream, true);
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateStream">The stream of the document template file.</param>
		///<param name="includeContent">Whether to copy the document template text content to document.</param>
		public void ApplyTemplate(Stream templateStream, bool includeContent)
		{
			Package templatePackage = Package.Open(templateStream);
			try
			{
				PackagePart documentPart = null;
				XDocument documentDoc = null;
				foreach (PackagePart packagePart in templatePackage.GetParts())
				{
					switch (packagePart.Uri.ToString())
					{
						case "/word/document.xml":
							documentPart = packagePart;
							using (XmlReader xr = XmlReader.Create(packagePart.GetStream(FileMode.Open, FileAccess.Read)))
							{
								documentDoc = XDocument.Load(xr);
							}
							break;
						case "/_rels/.rels":
							if (!this.package.PartExists(packagePart.Uri))
							{
								this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
							}
							PackagePart globalRelsPart = this.package.GetPart(packagePart.Uri);
							using (
							  StreamReader tr = new StreamReader(
								packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
							{
								using (
								  StreamWriter tw = new StreamWriter(
									new PackagePartStream(globalRelsPart.GetStream(FileMode.Create, FileAccess.Write)), Encoding.UTF8))
								{
									tw.Write(tr.ReadToEnd());
								}
							}
							break;
						case "/word/_rels/document.xml.rels":
							break;
						default:
							if (!this.package.PartExists(packagePart.Uri))
							{
								this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
							}
							Encoding packagePartEncoding = Encoding.Default;
							if (packagePart.Uri.ToString().EndsWith(".xml") || packagePart.Uri.ToString().EndsWith(".rels"))
							{
								packagePartEncoding = Encoding.UTF8;
							}
							PackagePart nativePart = this.package.GetPart(packagePart.Uri);
							using (
							  StreamReader tr = new StreamReader(
								packagePart.GetStream(FileMode.Open, FileAccess.Read), packagePartEncoding))
							{
								using (
								  StreamWriter tw = new StreamWriter(
									new PackagePartStream(nativePart.GetStream(FileMode.Create, FileAccess.Write)), tr.CurrentEncoding))
								{
									tw.Write(tr.ReadToEnd());
								}
							}
							break;
					}
				}
				if (documentPart != null)
				{
					string mainContentType = documentPart.ContentType.Replace("template.main", "document.main");
					if (this.package.PartExists(documentPart.Uri))
					{
						this.package.DeletePart(documentPart.Uri);
					}
					PackagePart documentNewPart = this.package.CreatePart(
					  documentPart.Uri, mainContentType, documentPart.CompressionOption);
					using (XmlWriter xw = XmlWriter.Create(new PackagePartStream(documentNewPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						documentDoc.WriteTo(xw);
					}
					foreach (PackageRelationship documentPartRel in documentPart.GetRelationships())
					{
						documentNewPart.CreateRelationship(
						  documentPartRel.TargetUri,
						  documentPartRel.TargetMode,
						  documentPartRel.RelationshipType,
						  documentPartRel.Id);
					}
					this.mainPart = documentNewPart;
					this.mainDoc = documentDoc;
					PopulateDocument(this, templatePackage);

					// DragonFire: I added next line and recovered ApplyTemplate method. 
					// I do it, becouse  PopulateDocument(...) writes into field "settingsPart" the part of Template's package 
					//  and after line "templatePackage.Close();" in finally, field "settingsPart" becomes not available and method "Save" throw an exception...
					// That's why I recreated settingsParts and unlinked it from Template's package =)
					settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);
				}
				if (!includeContent)
				{
					foreach (Paragraph paragraph in this.Paragraphs)
					{
						paragraph.Remove(false);
					}
				}
			}
			finally
			{
				this.package.Flush();
				var documentRelsPart = this.package.GetPart(new Uri("/word/_rels/document.xml.rels", UriKind.Relative));
				using (TextReader tr = new StreamReader(documentRelsPart.GetStream(FileMode.Open, FileAccess.Read)))
				{
					tr.Read();
				}
				templatePackage.Close();
				PopulateDocument(Document, package);
			}
		}

		/// <summary>
		/// Add an Image into this document from a fully qualified or relative filename.
		/// </summary>
		/// <param name="filename">The fully qualified or relative filename.</param>
		/// <returns>An Image file.</returns>
		/// <example>
		/// Add an Image into this document from a fully qualified filename.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Add an Image from a file.
		///     document.AddImage(@"C:\Example\Image.png");
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="AddImage(System.IO.Stream)"/>
		/// <seealso cref="Paragraph.InsertPicture"/>
		public Image AddImage(string filename)
		{
			string contentType = "";

			// The extension this file has will be taken to be its format.
			switch (Path.GetExtension(filename))
			{
				case ".tiff": contentType = "image/tif"; break;
				case ".tif": contentType = "image/tif"; break;
				case ".png": contentType = "image/png"; break;
				case ".bmp": contentType = "image/png"; break;
				case ".gif": contentType = "image/gif"; break;
				case ".jpg": contentType = "image/jpg"; break;
				case ".jpeg": contentType = "image/jpeg"; break;
				default: contentType = "image/jpg"; break;
			}

			return AddImage(filename as object, contentType);
		}

		/// <summary>
		/// Add an Image into this document from a Stream.
		/// </summary>
		/// <param name="stream">A Stream stream.</param>
		/// <returns>An Image file.</returns>
		/// <example>
		/// Add an Image into a document using a Stream. 
		/// <code>
		/// // Open a FileStream fs to an Image.
		/// using (FileStream fs = new FileStream(@"C:\Example\Image.jpg", FileMode.Open))
		/// {
		///     // Load a document.
		///     using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		///     {
		///         // Add an Image from a filestream fs.
		///         document.AddImage(fs);
		///
		///         // Save all changes made to this document.
		///         document.Save();
		///     }// Release this document from memory.
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AddImage(string)"/>
		/// <seealso cref="Paragraph.InsertPicture"/>
		public Image AddImage(Stream stream)
		{
			return AddImage(stream as object);
		}

		/// <summary>
		/// Adds a hyperlink to a document and creates a Paragraph which uses it.
		/// </summary>
		/// <param name="text">The text as displayed by the hyperlink.</param>
		/// <param name="uri">The hyperlink itself.</param>
		/// <returns>Returns a hyperlink that can be inserted into a Paragraph.</returns>
		/// <example>
		/// Adds a hyperlink to a document and creates a Paragraph which uses it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Add a hyperlink to this document.
		///    Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));
		///    
		///    // Add a new Paragraph to this document.
		///    Paragraph p = document.InsertParagraph();
		///    p.Append("My favourite search engine is ");
		///    p.AppendHyperlink(h);
		///    p.Append(", I think it's great.");
		///
		///    // Save all changes made to this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Hyperlink AddHyperlink(string text, Uri uri)
		{
			XElement i = new XElement
			(
				XName.Get("hyperlink", DocX.w.NamespaceName),
				new XAttribute(r + "id", string.Empty),
				new XAttribute(w + "history", "1"),
				new XElement(XName.Get("r", DocX.w.NamespaceName),
				new XElement(XName.Get("rPr", DocX.w.NamespaceName),
				new XElement(XName.Get("rStyle", DocX.w.NamespaceName),
				new XAttribute(w + "val", "Hyperlink"))),
				new XElement(XName.Get("t", DocX.w.NamespaceName), text))
			);

			Hyperlink h = new Hyperlink(this, mainPart, i);

			h.text = text;
			h.uri = uri;

			AddHyperlinkStyleIfNotPresent();

			return h;
		}

		internal void AddHyperlinkStyleIfNotPresent()
		{
			Uri word_styles_Uri = new Uri("/word/styles.xml", UriKind.Relative);

			// If the internal document contains no /word/styles.xml create one.
			if (!package.PartExists(word_styles_Uri))
				HelperFunctions.AddDefaultStylesXml(package);

			// Load the styles.xml into memory.
			XDocument word_styles;
			using (TextReader tr = new StreamReader(package.GetPart(word_styles_Uri).GetStream()))
				word_styles = XDocument.Load(tr);

			bool hyperlinkStyleExists =
			(
				from s in word_styles.Element(w + "styles").Elements()
				let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
				where (styleId != null && styleId.Value == "Hyperlink")
				select s
			).Count() > 0;

			if (!hyperlinkStyleExists)
			{
				XElement style = new XElement
				(
					w + "style",
					new XAttribute(w + "type", "character"),
					new XAttribute(w + "styleId", "Hyperlink"),
						new XElement(w + "name", new XAttribute(w + "val", "Hyperlink")),
						new XElement(w + "basedOn", new XAttribute(w + "val", "DefaultParagraphFont")),
						new XElement(w + "uiPriority", new XAttribute(w + "val", "99")),
						new XElement(w + "unhideWhenUsed"),
						new XElement(w + "rsid", new XAttribute(w + "val", "0005416C")),
						new XElement
						(
							w + "rPr",
							new XElement(w + "color", new XAttribute(w + "val", "0000FF"), new XAttribute(w + "themeColor", "hyperlink")),
							new XElement
							(
								w + "u",
								new XAttribute(w + "val", "single")
							)
						)
				);
				word_styles.Element(w + "styles").Add(style);

				// Save the styles document.
				using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(word_styles_Uri).GetStream())))
					word_styles.Save(tw);
			}
		}

		private string GetNextFreeRelationshipID()
		{
			int id = (
				 from r in mainPart.GetRelationships()
				 where r.Id.Substring(0, 3).Equals("rId")
				 select int.Parse(r.Id.Substring(3))
			 ).DefaultIfEmpty().Max();

			// The conventiom for ids is rid01, rid02, etc
			string newId = id.ToString();
			int result;
			if (int.TryParse(newId, out result))
				return ("rId" + (result + 1));
			else
			{
				String guid = String.Empty;
				do
				{
					guid = Guid.NewGuid().ToString();
				} while (Char.IsDigit(guid[0]));
				return guid;
			}
		}

		/// <summary>
		/// Adds three new Headers to this document. One for the first page, one for odd pages and one for even pages.
		/// </summary>
		/// <example>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add header support to this document.
		///     document.AddHeaders();
		///
		///     // Get a collection of all headers in this document.
		///     Headers headers = document.Headers;
		///
		///     // The header used for the first page of this document.
		///     Header first = headers.first;
		///
		///     // The header used for odd pages of this document.
		///     Header odd = headers.odd;
		///
		///     // The header used for even pages of this document.
		///     Header even = headers.even;
		///
		///     // Force the document to use a different header for first, odd and even pages.
		///     document.DifferentFirstPage = true;
		///     document.DifferentOddAndEvenPages = true;
		///
		///     // Content can be added to the Headers in the same manor that it would be added to the main document.
		///     Paragraph p = first.InsertParagraph();
		///     p.Append("This is the first pages header.");
		///
		///     // Save all changes to this document.
		///     document.Save();    
		/// }// Release this document from memory.
		/// </example>
		public void AddHeaders()
		{
			AddHeadersOrFooters(true);

			headers.odd = Document.GetHeaderByType("default");
			headers.even = Document.GetHeaderByType("even");
			headers.first = Document.GetHeaderByType("first");
		}

		/// <summary>
		/// Adds three new Footers to this document. One for the first page, one for odd pages and one for even pages.
		/// </summary>
		/// <example>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add footer support to this document.
		///     document.AddFooters();
		///
		///     // Get a collection of all footers in this document.
		///     Footers footers = document.Footers;
		///
		///     // The footer used for the first page of this document.
		///     Footer first = footers.first;
		///
		///     // The footer used for odd pages of this document.
		///     Footer odd = footers.odd;
		///
		///     // The footer used for even pages of this document.
		///     Footer even = footers.even;
		///
		///     // Force the document to use a different footer for first, odd and even pages.
		///     document.DifferentFirstPage = true;
		///     document.DifferentOddAndEvenPages = true;
		///
		///     // Content can be added to the Footers in the same manor that it would be added to the main document.
		///     Paragraph p = first.InsertParagraph();
		///     p.Append("This is the first pages footer.");
		///
		///     // Save all changes to this document.
		///     document.Save();    
		/// }// Release this document from memory.
		/// </example>
		public void AddFooters()
		{
			AddHeadersOrFooters(false);

			footers.odd = Document.GetFooterByType("default");
			footers.even = Document.GetFooterByType("even");
			footers.first = Document.GetFooterByType("first");
		}

		/// <summary>
		/// Adds a Header to a document.
		/// If the document already contains a Header it will be replaced.
		/// </summary>
		/// <returns>The Header that was added to the document.</returns>
		internal void AddHeadersOrFooters(bool b)
		{
			string element = "ftr";
			string reference = "footer";
			if (b)
			{
				element = "hdr";
				reference = "header";
			}

			DeleteHeadersOrFooters(b);

			XElement sectPr = mainDoc.Root.Element(w + "body").Element(w + "sectPr");

			for (int i = 1; i < 4; i++)
			{
				string header_uri = string.Format("/word/{0}{1}.xml", reference, i);

				PackagePart headerPart = package.CreatePart(new Uri(header_uri, UriKind.Relative), string.Format("application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference), CompressionOption.Normal);
				PackageRelationship headerRelationship = mainPart.CreateRelationship(headerPart.Uri, TargetMode.Internal, string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference));

				XDocument header;

				// Load the document part into a XDocument object
				using (TextReader tr = new StreamReader(headerPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
				{
					header = XDocument.Parse
					(string.Format(@"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
                       <w:{0} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                         <w:p w:rsidR=""009D472B"" w:rsidRDefault=""009D472B"">
                           <w:pPr>
                             <w:pStyle w:val=""{1}"" />
                           </w:pPr>
                         </w:p>
                       </w:{0}>", element, reference)
					);
				}

				// Save the main document
				using (TextWriter tw = new StreamWriter(new PackagePartStream(headerPart.GetStream(FileMode.Create, FileAccess.Write))))
					header.Save(tw, SaveOptions.None);

				string type;
				switch (i)
				{
					case 1: type = "default"; break;
					case 2: type = "even"; break;
					case 3: type = "first"; break;
					default: throw new ArgumentOutOfRangeException();
				}

				sectPr.Add
				(
					new XElement
					(
						w + string.Format("{0}Reference", reference),
						new XAttribute(w + "type", type),
						new XAttribute(r + "id", headerRelationship.Id)
					)
				);
			}
		}

		internal void DeleteHeadersOrFooters(bool b)
		{
			string reference = "footer";
			if (b)
				reference = "header";

			// Get all header Relationships in this document.
			var header_relationships = mainPart.GetRelationshipsByType(string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference));

			foreach (PackageRelationship header_relationship in header_relationships)
			{
				// Get the TargetUri for this Part.
				Uri header_uri = header_relationship.TargetUri;

				// Check to see if the document actually contains the Part.
				if (!header_uri.OriginalString.StartsWith("/word/"))
					header_uri = new Uri("/word/" + header_uri.OriginalString, UriKind.Relative);

				if (package.PartExists(header_uri))
				{
					// Delete the Part
					package.DeletePart(header_uri);

					// Get all references to this Relationship in the document.
					var query =
					(
						from e in mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
						where (e.Name.LocalName == string.Format("{0}Reference", reference)) && (e.Attribute(r + "id").Value == header_relationship.Id)
						select e
					);

					// Remove all references to this Relationship in the document.
					for (int i = 0; i < query.Count(); i++)
						query.ElementAt(i).Remove();

					// Delete the Relationship.
					package.DeleteRelationship(header_relationship.Id);
				}
			}
		}

		internal Image AddImage(object o, string contentType = "image/jpeg")
		{
			// Open a Stream to the new image being added.
			Stream newImageStream;
			if (o is string)
				newImageStream = new FileStream(o as string, FileMode.Open, FileAccess.Read);
			else
				newImageStream = o as Stream;

			// Get all image parts in word\document.xml

			PackageRelationshipCollection relationshipsByImages = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			List<PackagePart> imageParts = relationshipsByImages.Select(ir => package.GetParts().FirstOrDefault(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString()))).Where(e => e != null).ToList();

			foreach (PackagePart relsPart in package.GetParts().Where(part => part.Uri.ToString().Contains("/word/")).Where(part => part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml")))
			{
				XDocument relsPartContent;
				using (TextReader tr = new StreamReader(relsPart.GetStream(FileMode.Open, FileAccess.Read)))
					relsPartContent = XDocument.Load(tr);

				IEnumerable<XElement> imageRelationships =
				relsPartContent.Root.Elements().Where
				(
					imageRel =>
					imageRel.Attribute(XName.Get("Type")).Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				);

				foreach (XElement imageRelationship in imageRelationships)
				{
					if (imageRelationship.Attribute(XName.Get("Target")) != null)
					{
						string targetMode = string.Empty;

						XAttribute targetModeAttibute = imageRelationship.Attribute(XName.Get("TargetMode"));
						if (targetModeAttibute != null)
						{
							targetMode = targetModeAttibute.Value;
						}

						if (!targetMode.Equals("External"))
						{
							string imagePartUri = Path.Combine(Path.GetDirectoryName(relsPart.Uri.ToString()), imageRelationship.Attribute(XName.Get("Target")).Value);
							imagePartUri = Path.GetFullPath(imagePartUri.Replace("\\_rels", string.Empty));
							imagePartUri = imagePartUri.Replace(Path.GetFullPath("\\"), string.Empty).Replace("\\", "/");

							if (!imagePartUri.StartsWith("/"))
								imagePartUri = "/" + imagePartUri;

							PackagePart imagePart = package.GetPart(new Uri(imagePartUri, UriKind.Relative));
							imageParts.Add(imagePart);
						}
					}
				}
			}

			// Loop through each image part in this document.
			foreach (PackagePart pp in imageParts)
			{
				// Open a tempory Stream to this image part.
				using (Stream tempStream = pp.GetStream(FileMode.Open, FileAccess.Read))
				{
					// Compare this image to the new image being added.
					if (HelperFunctions.IsSameFile(tempStream, newImageStream))
					{
						// Get the image object for this image part
						string id = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
						.Where(r => r.TargetUri == pp.Uri)
						.Select(r => r.Id).First();

						// Return the Image object
						return Images.Where(i => i.Id == id).First();
					}
				}
			}

			string imgPartUriPath = string.Empty;
			string extension = contentType.Substring(contentType.LastIndexOf("/") + 1);
			do
			{
				// Create a new image part.
				imgPartUriPath = string.Format
				(
					"/word/media/{0}.{1}",
					Guid.NewGuid().ToString(), // The unique part.
					extension
				);

			} while (package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)));

			// We are now guareenteed that imgPartUriPath is unique.
			PackagePart img = package.CreatePart(new Uri(imgPartUriPath, UriKind.Relative), contentType, CompressionOption.Normal);

			// Create a new image relationship
			PackageRelationship rel = mainPart.CreateRelationship(img.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

			// Open a Stream to the newly created Image part.
			using (Stream stream = new PackagePartStream(img.GetStream(FileMode.Create, FileAccess.Write)))
			{
				// Using the Stream to the real image, copy this streams data into the newly create Image part.
				using (newImageStream)
				{
					byte[] bytes = new byte[newImageStream.Length];
					newImageStream.Read(bytes, 0, (int)newImageStream.Length);
					stream.Write(bytes, 0, (int)newImageStream.Length);
				}// Close the Stream to the new image.
			}// Close the Stream to the new image part.

			return new Image(this, rel);
		}

		/// <summary>
		/// Save this document back to the location it was loaded from.
		/// </summary>
		/// <example>
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Add an Image from a file.
		///     document.AddImage(@"C:\Example\Image.jpg");
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="DocX.SaveAs(string)"/>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Load(string)"/> 
		/// <!-- 
		/// Bug found and fixed by krugs525 on August 12 2009.
		/// Use TFS compare to see exact code change.
		/// -->
		public void Save()
		{
			Headers headers = Headers;

			// Save the main document
			using (TextWriter tw = new StreamWriter(new PackagePartStream(mainPart.GetStream(FileMode.Create, FileAccess.Write))))
				mainDoc.Save(tw, SaveOptions.None);

			if (settings == null)
			{
				using (TextReader tr = new StreamReader(settingsPart.GetStream()))
					settings = XDocument.Load(tr);
			}

			XElement body = mainDoc.Root.Element(w + "body");
			XElement sectPr = body.Descendants(w + "sectPr").FirstOrDefault();

			if (sectPr != null)
			{
				var evenHeaderRef =
				(
					from e in mainDoc.Descendants(w + "headerReference")
					let type = e.Attribute(w + "type")
					where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
					select e.Attribute(r + "id").Value
				 ).LastOrDefault();

				if (evenHeaderRef != null)
				{
					XElement even = headers.even.Xml;

					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(evenHeaderRef).TargetUri
					);

					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							even
						).Save(tw, SaveOptions.None);
					}
				}

				var oddHeaderRef =
				(
					from e in mainDoc.Descendants(w + "headerReference")
					let type = e.Attribute(w + "type")
					where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
					select e.Attribute(r + "id").Value
				 ).LastOrDefault();

				if (oddHeaderRef != null)
				{
					XElement odd = headers.odd.Xml;

					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(oddHeaderRef).TargetUri
					);

					// Save header1
					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							odd
						).Save(tw, SaveOptions.None);
					}
				}

				var firstHeaderRef =
				(
					from e in mainDoc.Descendants(w + "headerReference")
					let type = e.Attribute(w + "type")
					where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
					select e.Attribute(r + "id").Value
				 ).LastOrDefault();

				if (firstHeaderRef != null)
				{
					XElement first = headers.first.Xml;
					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(firstHeaderRef).TargetUri
					);

					// Save header3
					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							first
						).Save(tw, SaveOptions.None);
					}
				}

				var oddFooterRef =
				(
					from e in mainDoc.Descendants(w + "footerReference")
					let type = e.Attribute(w + "type")
					where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
					select e.Attribute(r + "id").Value
				 ).LastOrDefault();

				if (oddFooterRef != null)
				{
					XElement odd = footers.odd.Xml;
					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(oddFooterRef).TargetUri
					);

					// Save header1
					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							odd
						).Save(tw, SaveOptions.None);
					}
				}

				var evenFooterRef =
				(
					from e in mainDoc.Descendants(w + "footerReference")
					let type = e.Attribute(w + "type")
					where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
					select e.Attribute(r + "id").Value
				 ).LastOrDefault();

				if (evenFooterRef != null)
				{
					XElement even = footers.even.Xml;
					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(evenFooterRef).TargetUri
					);

					// Save header2
					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							even
						).Save(tw, SaveOptions.None);
					}
				}

				var firstFooterRef =
				(
					 from e in mainDoc.Descendants(w + "footerReference")
					 let type = e.Attribute(w + "type")
					 where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
					 select e.Attribute(r + "id").Value
				).LastOrDefault();

				if (firstFooterRef != null)
				{
					XElement first = footers.first.Xml;
					Uri target = PackUriHelper.ResolvePartUri
					(
						mainPart.Uri,
						mainPart.GetRelationship(firstFooterRef).TargetUri
					);

					// Save header3
					using (TextWriter tw = new StreamWriter(new PackagePartStream(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument
						(
							new XDeclaration("1.0", "UTF-8", "yes"),
							first
						).Save(tw, SaveOptions.None);
					}
				}

				// Save the settings document.
				using (TextWriter tw = new StreamWriter(new PackagePartStream(settingsPart.GetStream(FileMode.Create, FileAccess.Write))))
					settings.Save(tw, SaveOptions.None);

				if (endnotesPart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(endnotesPart.GetStream(FileMode.Create, FileAccess.Write))))
						endnotes.Save(tw, SaveOptions.None);
				}

				if (footnotesPart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(footnotesPart.GetStream(FileMode.Create, FileAccess.Write))))
						footnotes.Save(tw, SaveOptions.None);
				}

				if (stylesPart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(stylesPart.GetStream(FileMode.Create, FileAccess.Write))))
						styles.Save(tw, SaveOptions.None);
				}

				if (stylesWithEffectsPart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(stylesWithEffectsPart.GetStream(FileMode.Create, FileAccess.Write))))
						stylesWithEffects.Save(tw, SaveOptions.None);
				}

				if (numberingPart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(numberingPart.GetStream(FileMode.Create, FileAccess.Write))))
						numbering.Save(tw, SaveOptions.None);
				}

				if (fontTablePart != null)
				{
					using (TextWriter tw = new StreamWriter(new PackagePartStream(fontTablePart.GetStream(FileMode.Create, FileAccess.Write))))
						fontTable.Save(tw, SaveOptions.None);
				}
			}

			// Close the document so that it can be saved.
			package.Flush();

			#region Save this document back to a file or stream, that was specified by the user at save time.
			if (filename != null)
			{
				using (FileStream fs = new FileStream(filename, FileMode.Create))
				{
					fs.Write(memoryStream.ToArray(), 0, (int)memoryStream.Length);
				}
			}
			else
			{
				if (stream.CanSeek) // 2013-05-25: Check if stream can be seeked to support System.Web.HttpResponseStream
				{
					// Set the length of this stream to 0
					stream.SetLength(0);

					// Write to the beginning of the stream
					stream.Position = 0;
				}

				memoryStream.WriteTo(stream);
				memoryStream.Flush();
			}
			#endregion
		}

		/// <summary>
		/// Save this document to a file.
		/// </summary>
		/// <param name="filename">The filename to save this document as.</param>
		/// <example>
		/// Load a document from one file and save it to another.
		/// <code>
		/// // Load a document using its fully qualified filename.
		/// DocX document = DocX.Load(@"C:\Example\Test1.docx");
		///
		/// // Insert a new Paragraph
		/// document.InsertParagraph("Hello world!", false);
		///
		/// // Save the document to a new location.
		/// document.SaveAs(@"C:\Example\Test2.docx");
		/// </code>
		/// </example>
		/// <example>
		/// Load a document from a Stream and save it to a file.
		/// <code>
		/// DocX document;
		/// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		/// {
		///     // Load a document using a stream.
		///     document = DocX.Load(fs1);
		///
		///     // Insert a new Paragraph
		///     document.InsertParagraph("Hello world again!", false);
		/// }
		///    
		/// // Save the document to a new location.
		/// document.SaveAs(@"C:\Example\Test2.docx");
		/// </code>
		/// </example>
		/// <seealso cref="DocX.Save()"/>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Load(string)"/>
		public void SaveAs(string filename)
		{
			this.filename = filename;
			this.stream = null;
			Save();
		}

		/// <summary>
		/// Save this document to a Stream.
		/// </summary>
		/// <param name="stream">The Stream to save this document to.</param>
		/// <example>
		/// Load a document from a file and save it to a Stream.
		/// <code>
		/// // Place holder for a document.
		/// DocX document;
		///
		/// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		/// {
		///     // Load a document using a stream.
		///     document = DocX.Load(fs1);
		///
		///     // Insert a new Paragraph
		///     document.InsertParagraph("Hello world again!", false);
		/// }
		///
		/// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
		/// {
		///     // Save the document to a different stream.
		///     document.SaveAs(fs2);
		/// }
		///
		/// // Release this document from memory.
		/// document.Dispose();
		/// </code>
		/// </example>
		/// <example>
		/// Load a document from one Stream and save it to another.
		/// <code>
		/// DocX document;
		/// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		/// {
		///     // Load a document using a stream.
		///     document = DocX.Load(fs1);
		///
		///     // Insert a new Paragraph
		///     document.InsertParagraph("Hello world again!", false);
		/// }
		/// 
		/// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
		/// {
		///     // Save the document to a different stream.
		///     document.SaveAs(fs2);
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="DocX.Save()"/>
		/// <seealso cref="DocX.Load(System.IO.Stream)"/>
		/// <seealso cref="DocX.Load(string)"/>
		public void SaveAs(Stream stream)
		{
			this.filename = null;
			this.stream = stream;
			Save();
		}

		/// <summary>
		/// Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
		/// </summary>
		///<param name="propertyName">The property name.</param>
		///<param name="propertyValue">The property value.</param>
		///<example>
		/// Add a core properties of each type to a document.
		/// <code>
		/// // Load Example.docx
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // If this document does not contain a core property called 'forename', create one.
		///     if (!document.CoreProperties.ContainsKey("forename"))
		///     {
		///         // Create a new core property called 'forename' and set its value.
		///         document.AddCoreProperty("forename", "Cathal");
		///     }
		///
		///     // Get this documents core property called 'forename'.
		///     string forenameValue = document.CoreProperties["forename"];
		///
		///     // Print all of the information about this core property to Console.
		///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", "forename", forenameValue));
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// } // Release this document from memory.
		///
		/// // Wait for the user to press a key before exiting.
		/// Console.ReadKey();
		/// </code>
		/// </example>
		/// <seealso cref="CoreProperties"/>
		/// <seealso cref="CustomProperty"/>
		/// <seealso cref="CustomProperties"/>
		public void AddCoreProperty(string propertyName, string propertyValue)
		{
			string propertyNamespacePrefix = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[0] : "cp";
			string propertyLocalName = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[1] : propertyName;

			// If this document does not contain a coreFilePropertyPart create one.)
			if (!package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
				throw new Exception("Core properties part doesn't exist.");

			XDocument corePropDoc;
			PackagePart corePropPart = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
			using (TextReader tr = new StreamReader(corePropPart.GetStream(FileMode.Open, FileAccess.Read)))
			{
				corePropDoc = XDocument.Load(tr);
			}

			XElement corePropElement =
			  (from propElement in corePropDoc.Root.Elements()
			   where (propElement.Name.LocalName.Equals(propertyLocalName))
			   select propElement).SingleOrDefault();
			if (corePropElement != null)
			{
				corePropElement.SetValue(propertyValue);
			}
			else
			{
				var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix);
				corePropDoc.Root.Add(new XElement(XName.Get(propertyLocalName, propertyNamespace.NamespaceName), propertyValue));
			}

			using (TextWriter tw = new StreamWriter(new PackagePartStream(corePropPart.GetStream(FileMode.Create, FileAccess.Write))))
			{
				corePropDoc.Save(tw);
			}
			UpdateCorePropertyValue(this, propertyLocalName, propertyValue);
		}

		internal static void UpdateCorePropertyValue(DocX document, string corePropertyName, string corePropertyValue)
		{
			string matchPattern = string.Format(@"(DOCPROPERTY)?{0}\\\*MERGEFORMAT", corePropertyName).ToLower();
			foreach (XElement e in document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
			{
				string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();

				if (Regex.IsMatch(attr_value, matchPattern))
				{
					XElement firstRun = e.Element(w + "r");
					XElement firstText = firstRun.Element(w + "t");
					XElement rPr = firstText.Element(w + "rPr");

					// Delete everything and insert updated text value
					e.RemoveNodes();

					XElement t = new XElement(w + "t", rPr, corePropertyValue);
					Novacode.Text.PreserveSpace(t);
					e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
				}
			}

			#region Headers

			IEnumerable<PackagePart> headerParts = from headerPart in document.package.GetParts()
												   where (Regex.IsMatch(headerPart.Uri.ToString(), @"/word/header\d?.xml"))
												   select headerPart;
			foreach (PackagePart pp in headerParts)
			{
				XDocument header = XDocument.Load(new StreamReader(pp.GetStream()));

				foreach (XElement e in header.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(attr_value, matchPattern))
					{
						XElement firstRun = e.Element(w + "r");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(w + "t", corePropertyValue);
						Novacode.Text.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
					}
				}

				using (TextWriter tw = new StreamWriter(new PackagePartStream(pp.GetStream(FileMode.Create, FileAccess.Write))))
					header.Save(tw);
			}
			#endregion

			#region Footers
			IEnumerable<PackagePart> footerParts = from footerPart in document.package.GetParts()
												   where (Regex.IsMatch(footerPart.Uri.ToString(), @"/word/footer\d?.xml"))
												   select footerPart;
			foreach (PackagePart pp in footerParts)
			{
				XDocument footer = XDocument.Load(new StreamReader(pp.GetStream()));

				foreach (XElement e in footer.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(attr_value, matchPattern))
					{
						XElement firstRun = e.Element(w + "r");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(w + "t", corePropertyValue);
						Novacode.Text.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
					}
				}

				using (TextWriter tw = new StreamWriter(new PackagePartStream(pp.GetStream(FileMode.Create, FileAccess.Write))))
					footer.Save(tw);
			}
			#endregion
			PopulateDocument(document, document.package);
		}

		/// <summary>
		/// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
		/// </summary>
		/// <param name="cp">The CustomProperty to add to this document.</param>
		/// <example>
		/// Add a custom properties of each type to a document.
		/// <code>
		/// // Load Example.docx
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // A CustomProperty called forename which stores a string.
		///     CustomProperty forename;
		///
		///     // If this document does not contain a custom property called 'forename', create one.
		///     if (!document.CustomProperties.ContainsKey("forename"))
		///     {
		///         // Create a new custom property called 'forename' and set its value.
		///         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
		///     }
		///
		///     // Get this documents custom property called 'forename'.
		///     forename = document.CustomProperties["forename"];
		///
		///     // Print all of the information about this CustomProperty to Console.
		///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// } // Release this document from memory.
		///
		/// // Wait for the user to press a key before exiting.
		/// Console.ReadKey();
		/// </code>
		/// </example>
		/// <seealso cref="CustomProperty"/>
		/// <seealso cref="CustomProperties"/>
		public void AddCustomProperty(CustomProperty cp)
		{
			// If this document does not contain a customFilePropertyPart create one.
			if (!package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
				HelperFunctions.CreateCustomPropertiesPart(this);

			XDocument customPropDoc;
			PackagePart customPropPart = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
			using (TextReader tr = new StreamReader(customPropPart.GetStream(FileMode.Open, FileAccess.Read)))
				customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

			// Each custom property has a PID, get the highest PID in this document.
			IEnumerable<int> pids =
			(
				from d in customPropDoc.Descendants()
				where d.Name.LocalName == "property"
				select int.Parse(d.Attribute(XName.Get("pid")).Value)
			);

			int pid = 1;
			if (pids.Count() > 0)
				pid = pids.Max();

			// Check if a custom property already exists with this name
			// 2013-05-25: IgnoreCase while searching for custom property as it would produce a currupted docx.
			var customProperty =
			(
				from d in customPropDoc.Descendants()
				where (d.Name.LocalName == "property") && (d.Attribute(XName.Get("name")).Value.Equals(cp.Name, StringComparison.InvariantCultureIgnoreCase))
				select d
			).SingleOrDefault();

			// If a custom property with this name already exists remove it.
			if (customProperty != null)
				customProperty.Remove();

			XElement propertiesElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
			propertiesElement.Add
			(
				new XElement
				(
					XName.Get("property", customPropertiesSchema.NamespaceName),
					new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
					new XAttribute("pid", pid + 1),
					new XAttribute("name", cp.Name),
						new XElement(customVTypesSchema + cp.Type, cp.Value ?? "")
				)
			);

			// Save the custom properties
			using (TextWriter tw = new StreamWriter(new PackagePartStream(customPropPart.GetStream(FileMode.Create, FileAccess.Write))))
				customPropDoc.Save(tw, SaveOptions.None);

			// Refresh all fields in this document which display this custom property.
			UpdateCustomPropertyValue(this, cp.Name, (cp.Value ?? "").ToString());
		}

		/// <summary>
		/// Update the custom properties inside the document
		/// </summary>
		/// <param name="document">The DocX document</param>
		/// <param name="customPropertyName">The property used inside the document</param>
		/// <param name="customPropertyValue">The new value for the property</param>
		/// <remarks>Different version of Word create different Document XML.</remarks>
		internal static void UpdateCustomPropertyValue(DocX document, string customPropertyName, string customPropertyValue)
		{
			// A list of documents, which will contain, The Main Document and if they exist: header1, header2, header3, footer1, footer2, footer3.
			List<XElement> documents = new List<XElement> { document.mainDoc.Root };

			// Check if each header exists and add if if so.
			#region Headers
			Headers headers = document.Headers;
			if (headers.first != null)
				documents.Add(headers.first.Xml);
			if (headers.odd != null)
				documents.Add(headers.odd.Xml);
			if (headers.even != null)
				documents.Add(headers.even.Xml);
			#endregion

			// Check if each footer exists and add if if so.
			#region Footers
			Footers footers = document.Footers;
			if (footers.first != null)
				documents.Add(footers.first.Xml);
			if (footers.odd != null)
				documents.Add(footers.odd.Xml);
			if (footers.even != null)
				documents.Add(footers.even.Xml);
			#endregion

			var matchCustomPropertyName = customPropertyName;
			if (customPropertyName.Contains(" ")) matchCustomPropertyName = "\"" + customPropertyName + "\"";
			string match_value = string.Format(@"DOCPROPERTY  {0}  \* MERGEFORMAT", matchCustomPropertyName).Replace(" ", string.Empty);

			// Process each document in the list.
			foreach (XElement doc in documents)
			{
				#region Word 2010+
				foreach (XElement e in doc.Descendants(XName.Get("instrText", w.NamespaceName)))
				{

					string attr_value = e.Value.Replace(" ", string.Empty).Trim();

					if (attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase))
					{
						XNode node = e.Parent.NextNode;
						bool found = false;
						while (true)
						{
							if (node.NodeType == XmlNodeType.Element)
							{
								var ele = node as XElement;
								var match = ele.Descendants(XName.Get("t", w.NamespaceName));
								if (match.Count() > 0)
								{
									if (!found)
									{
										match.First().Value = customPropertyValue;
										found = true;
									}
									else
									{
										ele.RemoveNodes();
									}
								}
								else
								{
									match = ele.Descendants(XName.Get("fldChar", w.NamespaceName));
									if (match.Count() > 0)
									{
										var endMatch = match.First().Attribute(XName.Get("fldCharType", w.NamespaceName));
										if (endMatch != null && endMatch.Value == "end")
										{
											break;
										}
									}
								}
							}
							node = node.NextNode;
						}
					}
				}
				#endregion

				#region < Word 2010
				foreach (XElement e in doc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim();

					if (attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase))
					{
						XElement firstRun = e.Element(w + "r");
						XElement firstText = firstRun.Element(w + "t");
						XElement rPr = firstText.Element(w + "rPr");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(w + "t", rPr, customPropertyValue);
						Novacode.Text.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
					}
				}
				#endregion
			}
		}

		/// <summary></summary>
		/// <returns></returns>
		public override Paragraph InsertParagraph()
		{
			Paragraph p = base.InsertParagraph();
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(index, p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text)
		{
			Paragraph p = base.InsertParagraph(text);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges, formatting);
			p.PackagePart = mainPart;

			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public Paragraph[] InsertParagraphs(string text)
		{
			string[] textArray = text.Split('\n');
			List<Paragraph> paragraphs = new List<Paragraph>();
			foreach (var textForParagraph in textArray)
			{
				Paragraph p = base.InsertParagraph(text);
				p.PackagePart = mainPart;
				paragraphs.Add(p);
			}
			return paragraphs.ToArray();
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Content> Contents
		{
			get
			{
				ReadOnlyCollection<Content> l = base.Contents;
				foreach (var content in l) content.PackagePart = mainPart;
				return l;
			}
		}

		/// <summary></summary>
		/// <param name="el"></param>
		public void SetContent(XElement el)
		{
			foreach (XElement e in el.Elements())
			{
				(from d in Document.Contents where d.Name == e.Name select d).First().SetText(e.Value);
			}
		}

		/// <summary></summary>
		/// <param name="dict"></param>
		public void SetContent(Dictionary<string, string> dict)
		{
			foreach (KeyValuePair<string, string> item in dict)
			{
				(from d in Document.Contents where d.Name == item.Key select d).First().SetText(item.Value);
			}
		}

		/// <summary></summary>
		/// <param name="path"></param>
		public void SetContent(string path)
		{
			XDocument doc = XDocument.Load(path);
			SetContent(doc);
		}

		/// <summary></summary>
		/// <param name="xmlDoc"></param>
		public void SetContent(XDocument xmlDoc)
		{
			foreach (XElement e in xmlDoc.ElementsAfterSelf())
			{
				(from d in Document.Contents where d.Name == e.Name select d).First().SetText(e.Value);
			}
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> l = base.Paragraphs;
				foreach (var paragraph in l) paragraph.PackagePart = mainPart;
				return l;
			}
		}

		/// <summary></summary>
		public override List<List> Lists
		{
			get
			{
				List<List> l = base.Lists;
				l.ForEach(x => x.Items.ForEach(i => i.PackagePart = mainPart));
				return l;
			}
		}

		/// <summary></summary>
		public override List<Table> Tables
		{
			get
			{
				List<Table> l = base.Tables;
				l.ForEach(x => x.mainPart = mainPart);
				return l;
			}
		}


		/// <summary>Create an equation and insert it in the new paragraph</summary>
		public override Paragraph InsertEquation(string equation)
		{
			Paragraph p = base.InsertEquation(equation);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>Insert a chart in document</summary>
		public void InsertChart(Chart chart)
		{
			// Create a new chart part uri
			string chartPartUriPath = string.Empty;
			int chartIndex = 1;
			do
			{
				chartPartUriPath = string.Format("/word/charts/chart{0}.xml", chartIndex);
				chartIndex++;
			}
			while (package.PartExists(new Uri(chartPartUriPath, UriKind.Relative)));
			// Create chart part
			PackagePart chartPackagePart = package.CreatePart(new Uri(chartPartUriPath, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal);
			// Create a new chart relationship
			string relID = GetNextFreeRelationshipID();
			PackageRelationship rel = mainPart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relID);
			// Save a chart info the chartPackagePart
			using (TextWriter tw = new StreamWriter(new PackagePartStream(chartPackagePart.GetStream(FileMode.Create, FileAccess.Write)))) chart.Xml.Save(tw);
			// Insert a new chart into a paragraph.
			Paragraph p = InsertParagraph();
			XElement chartElement = new XElement(
				XName.Get("r", DocX.w.NamespaceName),
				new XElement(
					XName.Get("drawing", DocX.w.NamespaceName),
					new XElement(
						XName.Get("inline", DocX.wp.NamespaceName),
						new XElement(XName.Get("extent", DocX.wp.NamespaceName), new XAttribute("cx", "5486400"), new XAttribute("cy", "3200400")),
						new XElement(XName.Get("effectExtent", DocX.wp.NamespaceName), new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "19050"), new XAttribute("b", "19050")),
						new XElement(XName.Get("docPr", DocX.wp.NamespaceName), new XAttribute("id", "1"), new XAttribute("name", "chart")),
						new XElement(
							XName.Get("graphic", DocX.a.NamespaceName),
							new XElement(
								XName.Get("graphicData", DocX.a.NamespaceName),
								new XAttribute("uri", DocX.c.NamespaceName),
								new XElement(
									XName.Get("chart", DocX.c.NamespaceName),
									new XAttribute(XName.Get("id", DocX.r.NamespaceName), relID)
								)
							)
						)
					)
			   ));
			p.Xml.Add(chartElement);
		}

		/// <summary>Inserts a default TOC into the current document. Title: Table of contents. Swithces will be: TOC \h \o '1-3' \u \z</summary>
		/// <returns>The inserted TableOfContents</returns>
		public TableOfContents InsertDefaultTableOfContents()
		{
			return InsertTableOfContents("Table of contents", TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U);
		}

		/// <summary>Inserts a TOC into the current document</summary>
		/// <param name="title">The title of the TOC</param>
		/// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
		/// <param name="headerStyle">Lets you set the style name of the TOC header</param>
		/// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
		/// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
		/// <returns>The inserted TableOfContents</returns>
		public TableOfContents InsertTableOfContents(string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
		{
			var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			Xml.Add(toc.Xml);
			return toc;
		}

		/// <summary>Inserts at TOC into the current document before the provided <paramref name="reference"/></summary>
		/// <param name="reference">The paragraph to use as reference</param>
		/// <param name="title">The title of the TOC</param>
		/// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
		/// <param name="headerStyle">Lets you set the style name of the TOC header</param>
		/// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
		/// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
		/// <returns>The inserted TableOfContents</returns>
		public TableOfContents InsertTableOfContents(Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
		{
			var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			reference.Xml.AddBeforeSelf(toc.Xml);
			return toc;
		}

		/// <summary>
		/// Releases all resources used by this document.
		/// </summary>
		/// <example>
		/// If you take advantage of the using keyword, Dispose() is automatically called for you.
		/// <code>
		/// // Load document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///      // The document is only in memory while in this scope.
		///
		/// }// Dispose() is automatically called at this point.
		/// </code>
		/// </example>
		/// <example>
		/// This example is equilivant to the one above example.
		/// <code>
		/// // Load document.
		/// DocX document = DocX.Load(@"C:\Example\Test.docx");
		/// 
		/// // Do something with the document here.
		///
		/// // Dispose of the document.
		/// document.Dispose();
		/// </code>
		/// </example>
		public void Dispose()
		{
			package.Close();
		}

	}

	/// <summary></summary>
	public static class ExtensionsHeadings
	{

		/// <summary></summary>
		/// <param name="paragraph"></param>
		/// <param name="headingType"></param>
		/// <returns></returns>
		public static Paragraph Heading(this Paragraph paragraph, HeadingType headingType)
		{
			string StyleName = headingType.EnumDescription();
			paragraph.StyleName = StyleName;
			return paragraph;
		}

		/// <summary></summary>
		/// <param name="enumValue"></param>
		/// <returns></returns>
		public static string EnumDescription(this Enum enumValue)
		{
			if (enumValue == null || enumValue.ToString() == "0") return string.Empty;
			FieldInfo enumInfo = enumValue.GetType().GetField(enumValue.ToString());
			DescriptionAttribute[] enumAttributes = (DescriptionAttribute[])enumInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
			if (enumAttributes.Length > 0) return enumAttributes[0].Description;
			return enumValue.ToString();
		}

		/// <summary>
		/// From: http://stackoverflow.com/questions/4108828/generic-extension-method-to-see-if-an-enum-contains-a-flag
		/// Check to see if a flags enumeration has a specific flag set.
		/// </summary>
		/// <param name="variable">Flags enumeration to check</param>
		/// <param name="value">Flag to check for</param>
		/// <returns></returns>
		public static bool HasFlag(this Enum variable, Enum value)
		{
			if (variable == null) return false;
			if (value == null) throw new ArgumentNullException(nameof(value));
			// Not as good as the .NET 4 version of this function, but should be good enough
			if (!Enum.IsDefined(variable.GetType(), value)) throw new ArgumentException(string.Format("Enumeration type mismatch.  The flag is of type '{0}', was expecting '{1}'.", value.GetType(), variable.GetType()));
			ulong num = Convert.ToUInt64(value);
			return ((Convert.ToUInt64(variable) & num) == num);
		}

	}

	/// <summary>All DocX types are derived from DocXElement. This class contains properties which every element of a DocX must contain.</summary>
	public abstract class DocXElement
	{

		internal PackagePart mainPart;

		/// <summary></summary>
		public PackagePart PackagePart
		{
			get
			{
				return mainPart;
			}
			set
			{
				mainPart = value;
			}
		}

		/// <summary>
		/// This is the actual Xml that gives this element substance. 
		/// For example, a Paragraphs Xml might look something like the following
		/// <p>
		///     <r>
		///         <t>Hello World!</t>
		///     </r>
		/// </p>
		/// </summary>
		public XElement Xml
		{
			get;
			set;
		}

		/// <summary>
		/// This is a reference to the DocX object that this element belongs to.
		/// Every DocX element is connected to a document.
		/// </summary>
		internal DocX Document
		{
			get;
			set;
		}

		/// <summary>
		/// Store both the document and xml so that they can be accessed by derived types.
		/// </summary>
		/// <param name="document">The document that this element belongs to.</param>
		/// <param name="xml">The Xml that gives this element substance</param>
		public DocXElement(DocX document, XElement xml)
		{
			Document = document;
			Xml = xml;
		}

	}

	/// <summary>This class provides functions for inserting new DocXElements before or after the current DocXElement. Only certain DocXElements can support these functions without creating invalid documents, at the moment these are Paragraphs and Table.</summary>
	public abstract class InsertBeforeOrAfter : DocXElement
	{

		/// <summary></summary>
		/// <param name="document"></param>
		/// <param name="xml"></param>
		public InsertBeforeOrAfter(DocX document, XElement xml) : base(document, xml)
		{
		}

		/// <summary></summary>
		public virtual void InsertPageBreakBeforeSelf()
		{
			XElement p = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName),
					new XElement
					(
						XName.Get("r", DocX.w.NamespaceName),
							new XElement
							(
								XName.Get("br", DocX.w.NamespaceName),
								new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page")
							)
					)
			);
			Xml.AddBeforeSelf(p);
		}

		/// <summary></summary>
		public virtual void InsertPageBreakAfterSelf()
		{
			XElement p = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName),
					new XElement
					(
						XName.Get("r", DocX.w.NamespaceName),
							new XElement
							(
								XName.Get("br", DocX.w.NamespaceName),
								new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page")
							)
					)
			);
			Xml.AddAfterSelf(p);
		}

		/// <summary></summary>
		/// <param name="p"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphBeforeSelf(Paragraph p)
		{
			Xml.AddBeforeSelf(p.Xml);
			XElement newlyInserted = Xml.ElementsBeforeSelf().First();
			p.Xml = newlyInserted;
			return p;
		}

		/// <summary></summary>
		/// <param name="p"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphAfterSelf(Paragraph p)
		{
			Xml.AddAfterSelf(p.Xml);
			XElement newlyInserted = Xml.ElementsAfterSelf().First();
			//Dmitchern
			if (this as Paragraph != null)
			{
				return new Paragraph(Document, newlyInserted, (this as Paragraph).endIndex);
			}
			else
			{
				p.Xml = newlyInserted; //IMPORTANT: I think we have return new paragraph in any case, but I dont know what to put as startIndex parameter into Paragraph constructor
				return p;
			}
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphBeforeSelf(string text)
		{
			return InsertParagraphBeforeSelf(text, false, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphAfterSelf(string text)
		{
			return InsertParagraphAfterSelf(text, false, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
		{
			return InsertParagraphBeforeSelf(text, trackChanges, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
		{
			return InsertParagraphAfterSelf(text, trackChanges, new Formatting());
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
		{
			XElement newParagraph = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
			);
			if (trackChanges) newParagraph = Paragraph.CreateEdit(EditType.ins, DateTime.Now, newParagraph);
			Xml.AddBeforeSelf(newParagraph);
			XElement newlyInserted = Xml.ElementsBeforeSelf().Last();
			Paragraph p = new Paragraph(Document, newlyInserted, -1);
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
		{
			XElement newParagraph = new XElement
			(
				XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
			);
			if (trackChanges) newParagraph = Paragraph.CreateEdit(EditType.ins, DateTime.Now, newParagraph);
			Xml.AddAfterSelf(newParagraph);
			XElement newlyInserted = Xml.ElementsAfterSelf().First();
			Paragraph p = new Paragraph(Document, newlyInserted, -1);
			return p;
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public virtual Table InsertTableAfterSelf(int rowCount, int columnCount)
		{
			XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
			Xml.AddAfterSelf(newTable);
			XElement newlyInserted = Xml.ElementsAfterSelf().First();
			return new Table(Document, newlyInserted) { mainPart = mainPart };
		}

		/// <summary></summary>
		/// <param name="t"></param>
		/// <returns></returns>
		public virtual Table InsertTableAfterSelf(Table t)
		{
			Xml.AddAfterSelf(t.Xml);
			XElement newlyInserted = Xml.ElementsAfterSelf().First();
			//Dmitchern
			//return new table, dont affect parameter table
			return new Table(Document, newlyInserted) { mainPart = mainPart };
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public virtual Table InsertTableBeforeSelf(int rowCount, int columnCount)
		{
			XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
			Xml.AddBeforeSelf(newTable);
			XElement newlyInserted = Xml.ElementsBeforeSelf().Last();
			return new Table(Document, newlyInserted) { mainPart = mainPart };
		}

		/// <summary></summary>
		/// <param name="t"></param>
		/// <returns></returns>
		public virtual Table InsertTableBeforeSelf(Table t)
		{
			Xml.AddBeforeSelf(t.Xml);
			XElement newlyInserted = Xml.ElementsBeforeSelf().Last();
			//Dmitchern
			//return new table, dont affect parameter table
			return new Table(Document, newlyInserted) { mainPart = mainPart };
		}

	}

	/// <summary></summary>
	public static class XmlTemplateBases
	{

		/// <summary></summary>
		public const string TocXmlBase = @"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:sdtPr><w:docPartObj><w:docPartGallery w:val='Table of Contents'/><w:docPartUnique/></w:docPartObj>\</w:sdtPr><w:sdtEndPr><w:rPr><w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/><w:color w:val='auto'/><w:sz w:val='22'/><w:szCs w:val='22'/><w:lang w:eastAsia='en-US'/></w:rPr></w:sdtEndPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val='{0}'/></w:pPr><w:r><w:t>{1}</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val='TOC1'/><w:tabs><w:tab w:val='right' w:leader='dot' w:pos='{2}'/></w:tabs><w:rPr><w:noProof/></w:rPr></w:pPr><w:r><w:fldChar w:fldCharType='begin' w:dirty='true'/></w:r><w:r><w:instrText xml:space='preserve'> {3} </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r></w:p><w:p><w:r><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:fldChar w:fldCharType='end'/></w:r></w:p></w:sdtContent></w:sdt>";

		/// <summary></summary>
		public const string TocHeadingStyleBase = @"<w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:name w:val='TOC Heading'/><w:basedOn w:val='Heading1'/><w:next w:val='Normal'/><w:uiPriority w:val='39'/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:rsid w:val='00E67AA6'/><w:pPr><w:outlineLvl w:val='9'/></w:pPr><w:rPr><w:lang w:eastAsia='nb-NO'/></w:rPr></w:style>";

		/// <summary></summary>
		public const string TocElementStyleBase = @"<w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:name w:val='{1}' /><w:basedOn w:val='Normal' /><w:next w:val='Normal' /><w:autoRedefine /><w:uiPriority w:val='39' /><w:unhideWhenUsed /><w:pPr><w:spacing w:after='100' /><w:ind w:left='440' /></w:pPr></w:style>";

		/// <summary></summary>
		public const string TocHyperLinkStyleBase = @"<w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:name w:val='Hyperlink' /><w:basedOn w:val='Normal' /><w:uiPriority w:val='99' /><w:unhideWhenUsed /><w:rPr><w:color w:val='0000FF' w:themeColor='hyperlink' /><w:u w:val='single' /></w:rPr></w:style>";

	}

	/// <summary>Represents a font family</summary>
	public sealed class Font
	{
		/// <summary>
		/// Initializes a new instance of <see cref="Font" />
		/// </summary>
		/// <param name="name">The name of the font family</param>
		public Font(string name)
		{
			if (string.IsNullOrEmpty(name))
			{
				throw new ArgumentNullException(nameof(name));
			}

			Name = name;
		}

		/// <summary>
		/// The name of the font family
		/// </summary>
		public string Name { get; private set; }

		/// <summary>
		/// Returns a string representation of an object
		/// </summary>
		/// <returns>The name of the font family</returns>
		public override string ToString()
		{
			return Name;
		}
	}

	/// <summary></summary>
	public class Footer : Container, IParagraphContainer
	{

		/// <summary></summary>
		public bool PageNumbers
		{
			get
			{
				return false;
			}
			set
			{
				XElement e = XElement.Parse
				(@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:sdtPr>
                      <w:id w:val='157571950' />
                      <w:docPartObj>
                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />
                        <w:docPartUnique />
                      </w:docPartObj>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                          <w:jc w:val='center' />
                        </w:pPr>
                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>
                          <w:r>
                            <w:rPr>
                              <w:noProof />
                            </w:rPr>
                            <w:t>1</w:t>
                          </w:r>
                        </w:fldSimple>
                      </w:p>
                    </w:sdtContent>
                  </w:sdt>"
			   );

				Xml.AddFirst(e);
			}
		}

		internal Footer(DocX document, XElement xml, PackagePart mainPart) : base(document, xml)
		{
			this.mainPart = mainPart;
		}

		/// <summary></summary>
		/// <returns></returns>
		public override Paragraph InsertParagraph()
		{
			Paragraph p = base.InsertParagraph();
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(index, p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text)
		{
			Paragraph p = base.InsertParagraph(text);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="equation"></param>
		/// <returns></returns>
		public override Paragraph InsertEquation(string equation)
		{
			Paragraph p = base.InsertEquation(equation);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> l = base.Paragraphs;
				foreach (var paragraph in l) paragraph.mainPart = mainPart;
				return l;
			}
		}

		/// <summary></summary>
		public override List<Table> Tables
		{
			get
			{
				List<Table> l = base.Tables;
				l.ForEach(x => x.mainPart = mainPart);
				return l;
			}
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public new Table InsertTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = base.InsertTable(rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="t"></param>
		/// <returns></returns>
		public new Table InsertTable(int index, Table t)
		{
			Table t2 = base.InsertTable(index, t);
			t2.mainPart = mainPart;
			return t2;
		}

		/// <summary></summary>
		/// <param name="t"></param>
		/// <returns></returns>
		public new Table InsertTable(Table t)
		{
			t = base.InsertTable(t);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public new Table InsertTable(int index, int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = base.InsertTable(index, rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

	}

	/// <summary></summary>
	public class Footers
	{

		internal Footers()
		{
		}

		/// <summary></summary>
		public Footer odd;

		/// <summary></summary>
		public Footer even;

		/// <summary></summary>
		public Footer first;

	}

	/// <summary></summary>
	public class FormattedText : IComparable
	{

		/// <summary></summary>
		public FormattedText()
		{
		}

		/// <summary></summary>
		public int index;

		/// <summary></summary>
		public string text;

		/// <summary></summary>
		public Formatting formatting;

		/// <summary></summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public int CompareTo(object obj)
		{
			FormattedText other = (FormattedText)obj;
			FormattedText tf = this;
			if (other.formatting == null || tf.formatting == null) return -1;
			return tf.formatting.CompareTo(other.formatting);
		}

	}

	/// <summary>A text formatting</summary>
	public class Formatting : IComparable
	{

		private XElement rPr;

		private bool? hidden;

		private bool? bold;

		private bool? italic;

		private StrikeThrough? strikethrough;

		private Script? script;

		private Highlight? highlight;

		private double? size;

		private Color? fontColor;

		private Color? underlineColor;

		private UnderlineStyle? underlineStyle;

		private Misc? misc;

		private CapsStyle? capsStyle;

		private Font fontFamily;

		private int? percentageScale;

		private int? kerning;

		private int? position;

		private double? spacing;

		private CultureInfo language;

		/// <summary>A text formatting</summary>
		public Formatting()
		{
			capsStyle = Novacode.CapsStyle.none;
			strikethrough = Novacode.StrikeThrough.none;
			script = Novacode.Script.none;
			highlight = Novacode.Highlight.none;
			underlineStyle = Novacode.UnderlineStyle.none;
			misc = Novacode.Misc.none;
			// Use current culture by default
			language = CultureInfo.CurrentCulture;
			rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));
		}

		/// <summary>Text language</summary>
		public CultureInfo Language
		{
			get
			{
				return language;
			}
			set
			{
				language = value;
			}
		}

		/// <summary>Returns a new identical instance of Formatting</summary>
		/// <returns></returns>
		public Formatting Clone()
		{
			Formatting newf = new Formatting();
			newf.Bold = bold;
			newf.CapsStyle = capsStyle;
			newf.FontColor = fontColor;
			newf.FontFamily = fontFamily;
			newf.Hidden = hidden;
			newf.Highlight = highlight;
			newf.Italic = italic;
			if (kerning.HasValue) { newf.Kerning = kerning; }
			newf.Language = language;
			newf.Misc = misc;
			if (percentageScale.HasValue) { newf.PercentageScale = percentageScale; }
			if (position.HasValue) { newf.Position = position; }
			newf.Script = script;
			if (size.HasValue) { newf.Size = size; }
			if (spacing.HasValue) { newf.Spacing = spacing; }
			newf.StrikeThrough = strikethrough;
			newf.UnderlineColor = underlineColor;
			newf.UnderlineStyle = underlineStyle;
			return newf;
		}

		/// <summary></summary>
		/// <param name="rPr"></param>
		/// <returns></returns>
		public static Formatting Parse(XElement rPr)
		{
			Formatting formatting = new Formatting();

			// Build up the Formatting object.
			foreach (XElement option in rPr.Elements())
			{
				switch (option.Name.LocalName)
				{
					case "lang":
						formatting.Language = new CultureInfo(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null) ??
							option.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), null) ??
							option.GetAttribute(XName.Get("bidi", DocX.w.NamespaceName)));
						break;
					case "spacing":
						formatting.Spacing = Double.Parse(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 20.0;
						break;
					case "position":
						formatting.Position = Int32.Parse(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
						break;
					case "kern":
						formatting.Position = Int32.Parse(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
						break;
					case "w":
						formatting.PercentageScale = Int32.Parse(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
						break;
					// <w:sz w:val="20"/><w:szCs w:val="20"/>
					case "sz":
						formatting.Size = Int32.Parse(
							option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
						break;


					case "rFonts":
						formatting.FontFamily =
							new Font(
								option.GetAttribute(XName.Get("cs", DocX.w.NamespaceName), null) ??
								option.GetAttribute(XName.Get("ascii", DocX.w.NamespaceName), null) ??
								option.GetAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), null) ??
								option.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName)));
						break;
					case "color":
						try
						{
							string color = option.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
							formatting.FontColor = System.Drawing.ColorTranslator.FromHtml(string.Format("#{0}", color));
						}
						catch { }
						break;
					case "vanish": formatting.hidden = true; break;
					case "b": formatting.Bold = true; break;
					case "i": formatting.Italic = true; break;
					case "u":
						formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
						break;
					default: break;
				}
			}


			return formatting;
		}

		internal XElement Xml
		{
			get
			{
				rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));

				if (language != null)
					rPr.Add(new XElement(XName.Get("lang", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), language.Name)));

				if (spacing.HasValue)
					rPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing.Value * 20)));

				if (position.HasValue)
					rPr.Add(new XElement(XName.Get("position", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), position.Value * 2)));

				if (kerning.HasValue)
					rPr.Add(new XElement(XName.Get("kern", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning.Value * 2)));

				if (percentageScale.HasValue)
					rPr.Add(new XElement(XName.Get("w", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale)));

				if (fontFamily != null)
				{
					rPr.Add
					(
						new XElement
						(
							XName.Get("rFonts", DocX.w.NamespaceName),
							new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name),
							new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name), // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
							new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name)    // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
						)
					);
				}

				if (hidden.HasValue && hidden.Value)
					rPr.Add(new XElement(XName.Get("vanish", DocX.w.NamespaceName)));

				if (bold.HasValue && bold.Value)
					rPr.Add(new XElement(XName.Get("b", DocX.w.NamespaceName)));

				if (italic.HasValue && italic.Value)
					rPr.Add(new XElement(XName.Get("i", DocX.w.NamespaceName)));

				if (underlineStyle.HasValue)
				{
					switch (underlineStyle)
					{
						case Novacode.UnderlineStyle.none:
							break;
						case Novacode.UnderlineStyle.singleLine:
							rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
							break;
						case Novacode.UnderlineStyle.doubleLine:
							rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "double")));
							break;
						default:
							rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), underlineStyle.ToString())));
							break;
					}
				}

				if (underlineColor.HasValue)
				{
					// If an underlineColor has been set but no underlineStyle has been set
					if (underlineStyle == Novacode.UnderlineStyle.none)
					{
						// Set the underlineStyle to the default
						underlineStyle = Novacode.UnderlineStyle.singleLine;
						rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
					}

					rPr.Element(XName.Get("u", DocX.w.NamespaceName)).Add(new XAttribute(XName.Get("color", DocX.w.NamespaceName), underlineColor.Value.ToHex()));
				}

				if (strikethrough.HasValue)
				{
					switch (strikethrough)
					{
						case Novacode.StrikeThrough.none:
							break;
						case Novacode.StrikeThrough.strike:
							rPr.Add(new XElement(XName.Get("strike", DocX.w.NamespaceName)));
							break;
						case Novacode.StrikeThrough.doubleStrike:
							rPr.Add(new XElement(XName.Get("dstrike", DocX.w.NamespaceName)));
							break;
						default:
							break;
					}
				}

				if (script.HasValue)
				{
					switch (script)
					{
						case Novacode.Script.none:
							break;
						default:
							rPr.Add(new XElement(XName.Get("vertAlign", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), script.ToString())));
							break;
					}
				}

				if (size.HasValue)
				{
					rPr.Add(new XElement(XName.Get("sz", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2).ToString())));
					rPr.Add(new XElement(XName.Get("szCs", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2).ToString())));
				}

				if (fontColor.HasValue)
					rPr.Add(new XElement(XName.Get("color", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontColor.Value.ToHex())));

				if (highlight.HasValue)
				{
					switch (highlight)
					{
						case Novacode.Highlight.none:
							break;
						default:
							rPr.Add(new XElement(XName.Get("highlight", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight.ToString())));
							break;
					}
				}

				if (capsStyle.HasValue)
				{
					switch (capsStyle)
					{
						case Novacode.CapsStyle.none:
							break;
						default:
							rPr.Add(new XElement(XName.Get(capsStyle.ToString(), DocX.w.NamespaceName)));
							break;
					}
				}

				if (misc.HasValue)
				{
					switch (misc)
					{
						case Novacode.Misc.none:
							break;
						case Novacode.Misc.outlineShadow:
							rPr.Add(new XElement(XName.Get("outline", DocX.w.NamespaceName)));
							rPr.Add(new XElement(XName.Get("shadow", DocX.w.NamespaceName)));
							break;
						case Novacode.Misc.engrave:
							rPr.Add(new XElement(XName.Get("imprint", DocX.w.NamespaceName)));
							break;
						default:
							rPr.Add(new XElement(XName.Get(misc.ToString(), DocX.w.NamespaceName)));
							break;
					}
				}

				return rPr;
			}
		}

		/// <summary>
		/// This formatting will apply Bold.
		/// </summary>
		public bool? Bold { get { return bold; } set { bold = value; } }

		/// <summary>
		/// This formatting will apply Italic.
		/// </summary>
		public bool? Italic { get { return italic; } set { italic = value; } }

		/// <summary>
		/// This formatting will apply StrickThrough.
		/// </summary>
		public StrikeThrough? StrikeThrough { get { return strikethrough; } set { strikethrough = value; } }

		/// <summary>
		/// The script that this formatting should be, normal, superscript or subscript.
		/// </summary>
		public Script? Script { get { return script; } set { script = value; } }

		/// <summary>
		/// The Size of this text, must be between 0 and 1638.
		/// </summary>
		public double? Size
		{
			get { return size; }

			set
			{
				double? temp = value * 2;

				if (temp - (int)temp == 0)
				{
					if (value > 0 && value < 1639)
						size = value;
					else
						throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
				}

				else
					throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");
			}
		}

		/// <summary>
		/// Percentage scale must be one of the following values 200, 150, 100, 90, 80, 66, 50 or 33.
		/// </summary>
		public int? PercentageScale
		{
			get { return percentageScale; }

			set
			{
				if ((new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(value))
					percentageScale = value;
				else
					throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
			}
		}

		/// <summary>
		/// The Kerning to apply to this text must be one of the following values 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72.
		/// </summary>
		public int? Kerning
		{
			get { return kerning; }

			set
			{
				if (new int?[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains(value))
					kerning = value;
				else
					throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
			}
		}

		/// <summary>
		/// Text position must be in the range (-1585 - 1585).
		/// </summary>
		public int? Position
		{
			get { return position; }

			set
			{
				if (value > -1585 && value < 1585)
					position = value;
				else
					throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
			}
		}

		/// <summary>
		/// Text spacing must be in the range (-1585 - 1585).
		/// </summary>
		public double? Spacing
		{
			get { return spacing; }

			set
			{
				double? temp = value * 20;

				if (temp - (int)temp == 0)
				{
					if (value > -1585 && value < 1585)
						spacing = value;
					else
						throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
				}

				else
					throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");
			}
		}

		/// <summary>
		/// The colour of the text.
		/// </summary>
		public Color? FontColor { get { return fontColor; } set { fontColor = value; } }

		/// <summary>
		/// Highlight colour.
		/// </summary>
		public Highlight? Highlight { get { return highlight; } set { highlight = value; } }

		/// <summary>
		/// The Underline style that this formatting applies.
		/// </summary>
		public UnderlineStyle? UnderlineStyle { get { return underlineStyle; } set { underlineStyle = value; } }

		/// <summary>
		/// The underline colour.
		/// </summary>
		public Color? UnderlineColor { get { return underlineColor; } set { underlineColor = value; } }

		/// <summary>
		/// Misc settings.
		/// </summary>
		public Misc? Misc { get { return misc; } set { misc = value; } }

		/// <summary>
		/// Is this text hidden or visible.
		/// </summary>
		public bool? Hidden { get { return hidden; } set { hidden = value; } }

		/// <summary>
		/// Capitalization style.
		/// </summary>
		public CapsStyle? CapsStyle { get { return capsStyle; } set { capsStyle = value; } }

		/// <summary>
		/// The font family of this formatting.
		/// </summary>
		/// <!-- 
		/// Bug found and fixed by krugs525 on August 12 2009.
		/// Use TFS compare to see exact code change.
		/// -->
		public Font FontFamily { get { return fontFamily; } set { fontFamily = value; } }

		/// <summary></summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public int CompareTo(object obj)
		{
			Formatting other = (Formatting)obj;
			if (other.hidden != this.hidden) return -1;
			if (other.bold != this.bold) return -1;
			if (other.italic != this.italic) return -1;
			if (other.strikethrough != this.strikethrough) return -1;
			if (other.script != this.script) return -1;
			if (other.highlight != this.highlight) return -1;
			if (other.size != this.size) return -1;
			if (other.fontColor != this.fontColor) return -1;
			if (other.underlineColor != this.underlineColor) return -1;
			if (other.underlineStyle != this.underlineStyle) return -1;
			if (other.misc != this.misc) return -1;
			if (other.capsStyle != this.capsStyle) return -1;
			if (other.fontFamily != this.fontFamily) return -1;
			if (other.percentageScale != this.percentageScale) return -1;
			if (other.kerning != this.kerning) return -1;
			if (other.position != this.position) return -1;
			if (other.spacing != this.spacing) return -1;
			if (!other.language.Equals(this.language)) return -1;
			return 0;
		}

	}

	/// <summary></summary>
	public class Header : Container, IParagraphContainer
	{

		/// <summary></summary>
		public bool PageNumbers
		{
			get
			{
				return false;
			}
			set
			{
				XElement e = XElement.Parse
				(@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:sdtPr>
                      <w:id w:val='157571950' />
                      <w:docPartObj>
                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />
                        <w:docPartUnique />
                      </w:docPartObj>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                          <w:jc w:val='center' />
                        </w:pPr>
                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>
                          <w:r>
                            <w:rPr>
                              <w:noProof />
                            </w:rPr>
                            <w:t>1</w:t>
                          </w:r>
                        </w:fldSimple>
                      </w:p>
                    </w:sdtContent>
                  </w:sdt>"
			   );
				Xml.AddFirst(e);
				PageNumberParagraph = new Paragraph(Document, e.Descendants(XName.Get("p", DocX.w.NamespaceName)).SingleOrDefault(), 0);
			}
		}

		/// <summary></summary>
		public Paragraph PageNumberParagraph;

		internal Header(DocX document, XElement xml, PackagePart mainPart) : base(document, xml)
		{
			this.mainPart = mainPart;
		}

		/// <summary></summary>
		/// <returns></returns>
		public override Paragraph InsertParagraph()
		{
			Paragraph p = base.InsertParagraph();
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="p"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(index, p);
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(index, text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text)
		{
			Paragraph p = base.InsertParagraph(text);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="text"></param>
		/// <param name="trackChanges"></param>
		/// <param name="formatting"></param>
		/// <returns></returns>
		public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraph(text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		/// <param name="equation"></param>
		/// <returns></returns>
		public override Paragraph InsertEquation(string equation)
		{
			Paragraph p = base.InsertEquation(equation);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> l = base.Paragraphs;
				foreach (var paragraph in l) paragraph.mainPart = mainPart;
				return l;
			}
		}

		/// <summary></summary>
		public override List<Table> Tables
		{
			get
			{
				List<Table> l = base.Tables;
				l.ForEach(x => x.mainPart = mainPart);
				return l;
			}
		}

		/// <summary></summary>
		public List<Image> Images
		{
			get
			{
				PackageRelationshipCollection imageRelationships = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				if (imageRelationships.Count() > 0)
				{
					return
					(
						from i in imageRelationships
						select new Image(Document, i)
					).ToList();
				}

				return new List<Image>();
			}
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public new Table InsertTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = base.InsertTable(rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="t"></param>
		/// <returns></returns>
		public new Table InsertTable(int index, Table t)
		{
			Table t2 = base.InsertTable(index, t);
			t2.mainPart = mainPart;
			return t2;
		}

		/// <summary></summary>
		/// <param name="t"></param>
		/// <returns></returns>
		public new Table InsertTable(Table t)
		{
			t = base.InsertTable(t);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public new Table InsertTable(int index, int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1) throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			Table t = base.InsertTable(index, rowCount, columnCount);
			t.mainPart = mainPart;
			return t;
		}

	}

	/// <summary></summary>
	public class Headers
	{

		internal Headers()
		{
		}

		/// <summary></summary>
		public Header odd;

		/// <summary></summary>
		public Header even;

		/// <summary></summary>
		public Header first;

	}

	/// <summary>Represents a Hyperlink in a document</summary>
	public class Hyperlink : DocXElement
	{

		internal Uri uri;

		internal String text;

		internal Dictionary<PackagePart, PackageRelationship> hyperlink_rels;
		internal int type;
		internal String id;
		internal XElement instrText;
		internal List<XElement> runs;

		/// <summary>
		/// Remove a Hyperlink from this Paragraph only.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Add a hyperlink to this document.
		///    Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
		///
		///    // Add a Paragraph to this document and insert the hyperlink
		///    Paragraph p1 = document.InsertParagraph();
		///    p1.Append("This is a cool ").AppendHyperlink(h).Append(" .");
		///
		///    /* 
		///     * Remove the hyperlink from this Paragraph only. 
		///     * Note a reference to the hyperlink will still exist in the document and it can thus be reused.
		///     */
		///    p1.Hyperlinks[0].Remove();
		///
		///    // Add a new Paragraph to this document and reuse the hyperlink h.
		///    Paragraph p2 = document.InsertParagraph();
		///    p2.Append("This is the same cool ").AppendHyperlink(h).Append(" .");
		///
		///    document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void Remove()
		{
			Xml.Remove();
		}

		/// <summary>
		/// Change the Text of a Hyperlink.
		/// </summary>
		/// <example>
		/// Change the Text of a Hyperlink.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///    // Get all of the hyperlinks in this document
		///    List&lt;Hyperlink&gt; hyperlinks = document.Hyperlinks;
		///    
		///    // Change the first hyperlinks text and Uri
		///    Hyperlink h0 = hyperlinks[0];
		///    h0.Text = "DocX";
		///    h0.Uri = new Uri("http://docx.codeplex.com");
		///
		///    // Save this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public string Text
		{
			get
			{
				return this.text;
			}

			set
			{
				XElement rPr =
					new XElement
					(
						DocX.w + "rPr",
						new XElement
						(
							DocX.w + "rStyle",
							new XAttribute(DocX.w + "val", "Hyperlink")
						)
					);

				// Format and add the new text.
				List<XElement> newRuns = HelperFunctions.FormatInput(value, rPr);

				if (type == 0)
				{
					// Get all the runs in this Text.
					var runs = from r in Xml.Elements()
							   where r.Name.LocalName == "r"
							   select r;

					// Remove each run.
					for (int i = 0; i < runs.Count(); i++)
						runs.Remove();

					Xml.Add(newRuns);
				}

				else
				{
					XElement separate = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='separate'/> 
                    </w:r>");

					XElement end = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='end' /> 
                    </w:r>");

					runs.Last().AddAfterSelf(separate, newRuns, end);
					runs.ForEach(r => r.Remove());
				}

				this.text = value;
			}
		}

		/// <summary>
		/// Change the Uri of a Hyperlink.
		/// </summary>
		/// <example>
		/// Change the Uri of a Hyperlink.
		/// <code>
		/// <![CDATA[
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///    // Get all of the hyperlinks in this document
		///    List<Hyperlink> hyperlinks = document.Hyperlinks;
		///    
		///    // Change the first hyperlinks text and Uri
		///    Hyperlink h0 = hyperlinks[0];
		///    h0.Text = "DocX";
		///    h0.Uri = new Uri("http://docx.codeplex.com");
		///
		///    // Save this document.
		///    document.Save();
		/// }
		/// ]]>
		/// </code>
		/// </example>
		public Uri Uri
		{
			get
			{
				if (type == 0 && id != String.Empty)
				{
					PackageRelationship r = mainPart.GetRelationship(id);
					return r.TargetUri;
				}

				return this.uri;
			}

			set
			{
				if (type == 0)
				{
					PackageRelationship r = mainPart.GetRelationship(id);

					// Get all of the information about this relationship.
					TargetMode r_tm = r.TargetMode;
					string r_rt = r.RelationshipType;
					string r_id = r.Id;

					// Delete the relationship
					mainPart.DeleteRelationship(r_id);
					mainPart.CreateRelationship(value, r_tm, r_rt, r_id);
				}

				else
				{
					instrText.Value = "HYPERLINK " + "\"" + value + "\"";
				}

				this.uri = value;
			}
		}

		internal Hyperlink(DocX document, PackagePart mainPart, XElement i) : base(document, i)
		{
			this.type = 0;
			this.id = i.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;

			StringBuilder sb = new StringBuilder();
			HelperFunctions.GetTextRecursive(i, ref sb);
			this.text = sb.ToString();
		}

		internal Hyperlink(DocX document, XElement instrText, List<XElement> runs) : base(document, null)
		{
			this.type = 1;
			this.instrText = instrText;
			this.runs = runs;

			try
			{
				int start = instrText.Value.IndexOf("HYPERLINK \"") + "HYPERLINK \"".Length;
				int end = instrText.Value.IndexOf("\"", start);
				if (start != -1 && end != -1)
				{
					this.uri = new Uri(instrText.Value.Substring(start, end - start), UriKind.Absolute);

					StringBuilder sb = new StringBuilder();
					HelperFunctions.GetTextRecursive(new XElement(XName.Get("temp", DocX.w.NamespaceName), runs), ref sb);
					this.text = sb.ToString();
				}
			}

			catch (Exception e) { throw e; }
		}
	}

	/// <summary>Represents an Image embedded in a document</summary>
	public class Image
	{
		/// <summary>
		/// A unique id which identifies this Image.
		/// </summary>
		private string id;
		private DocX document;
		internal PackageRelationship pr;

		/// <summary></summary>
		/// <param name="mode"></param>
		/// <param name="access"></param>
		/// <returns></returns>
		public Stream GetStream(FileMode mode, FileAccess access)
		{
			string temp = pr.SourceUri.OriginalString;
			string start = temp.Remove(temp.LastIndexOf('/'));
			string end = pr.TargetUri.OriginalString;
			string full = start + "/" + end;
			return (new PackagePartStream(document.package.GetPart(new Uri(full, UriKind.Relative)).GetStream(mode, access)));
		}

		/// <summary>Returns the id of this Image</summary>
		public string Id
		{
			get
			{
				return id;
			}
		}

		internal Image(DocX document, PackageRelationship pr)
		{
			this.document = document;
			this.pr = pr;
			this.id = pr.Id;
		}

		/// <summary>
		/// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		/// </summary>
		/// <returns></returns>
		/// <example>
		/// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		/// <code>
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///    // Add an image to the document. 
		///    Image     i = document.AddImage(@"Image.jpg");
		///    
		///    // Create a picture i.e. (A custom view of an image)
		///    Picture   p = i.CreatePicture();
		///    p.FlipHorizontal = true;
		///    p.Rotation = 10;
		///
		///    // Create a new Paragraph.
		///    Paragraph par = document.InsertParagraph();
		///    
		///    // Append content to the Paragraph.
		///    par.Append("Here is a cool picture")
		///       .AppendPicture(p)
		///       .Append(" don't you think so?");
		///
		///    // Save all changes made to this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Picture CreatePicture()
		{
			return Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
		}

		/// <summary></summary>
		/// <param name="height"></param>
		/// <param name="width"></param>
		/// <returns></returns>
		public Picture CreatePicture(int height, int width)
		{
			Picture picture = Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
			picture.Height = height;
			picture.Width = width;
			return picture;
		}

		///<summary>Returns the name of the image file</summary>
		public string FileName
		{
			get
			{
				return Path.GetFileName(this.pr.TargetUri.ToString());
			}
		}

	}

	/// <summary>Represents a List in a document</summary>
	public class List : InsertBeforeOrAfter
	{

		/// <summary>
		/// This is a list of paragraphs that will be added to the document
		/// when the list is inserted into the document.
		/// The paragraph needs a numPr defined to be in this items collection.
		/// </summary>
		public List<Paragraph> Items { get; private set; }

		/// <summary>
		/// The numId used to reference the list settings in the numbering.xml
		/// </summary>
		public int NumId { get; private set; }

		/// <summary>
		/// The ListItemType (bullet or numbered) of the list.
		/// </summary>
		public ListItemType? ListType { get; private set; }

		internal List(DocX document, XElement xml) : base(document, xml)
		{
			Items = new List<Paragraph>();
			ListType = null;
		}

		/// <summary>
		/// Adds an item to the list.
		/// </summary>
		/// <param name="paragraph"></param>
		/// <exception cref="InvalidOperationException">
		/// Throws an InvalidOperationException if the item cannot be added to the list.
		/// </exception>
		public void AddItem(Paragraph paragraph)
		{
			if (paragraph.IsListItem)
			{
				var numIdNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "numId");
				var numId = Int32.Parse(numIdNode.Attribute(DocX.w + "val").Value);

				if (CanAddListItem(paragraph))
				{
					NumId = numId;
					Items.Add(paragraph);
				}
				else
					throw new InvalidOperationException("New list items can only be added to this list if they are have the same numId.");
			}
		}

		/// <summary></summary>
		/// <param name="paragraph"></param>
		/// <param name="start"></param>
		public void AddItemWithStartValue(Paragraph paragraph, int start)
		{
			//TODO: Update the numbering
			UpdateNumberingForLevelStartNumber(int.Parse(paragraph.IndentLevel.ToString()), start);
			if (ContainsLevel(start)) throw new InvalidOperationException("Cannot add a paragraph with a start value if another element already exists in this list with that level.");
			AddItem(paragraph);
		}

		private void UpdateNumberingForLevelStartNumber(int iLevel, int start)
		{
			var abstractNum = GetAbstractNum(NumId);
			var level = abstractNum.Descendants().First(el => el.Name.LocalName == "lvl" && el.GetAttribute(DocX.w + "ilvl") == iLevel.ToString());
			level.Descendants().First(el => el.Name.LocalName == "start").SetAttributeValue(DocX.w + "val", start);
		}

		/// <summary>
		/// Determine if it is able to add the item to the list
		/// </summary>
		/// <param name="paragraph"></param>
		/// <returns>
		/// Return true if AddItem(...) will succeed with the given paragraph.
		/// </returns>
		public bool CanAddListItem(Paragraph paragraph)
		{
			if (paragraph.IsListItem)
			{
				//var lvlNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "ilvl");
				var numIdNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "numId");
				var numId = Int32.Parse(numIdNode.Attribute(DocX.w + "val").Value);

				//Level = Int32.Parse(lvlNode.Attribute(DocX.w + "val").Value);
				if (NumId == 0 || (numId == NumId && numId > 0))
				{
					return true;
				}
			}
			return false;
		}

		/// <summary></summary>
		/// <param name="ilvl"></param>
		/// <returns></returns>
		public bool ContainsLevel(int ilvl)
		{
			return Items.Any(i => i.ParagraphNumberProperties.Descendants().First(el => el.Name.LocalName == "ilvl").Value == ilvl.ToString());
		}

		internal void CreateNewNumberingNumId(int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool continueNumbering = false)
		{
			ValidateDocXNumberingPartExists();
			if (Document.numbering.Root == null)
			{
				throw new InvalidOperationException("Numbering section did not instantiate properly.");
			}

			ListType = listType;

			var numId = GetMaxNumId() + 1;
			var abstractNumId = GetMaxAbstractNumId() + 1;

			XDocument listTemplate;
			switch (listType)
			{
				case ListItemType.Bulleted:
					listTemplate = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_bullet_abstract.xml.gz");
					break;
				case ListItemType.Numbered:
					listTemplate = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_decimal_abstract.xml.gz");
					break;
				default:
					throw new InvalidOperationException(string.Format("Unable to deal with ListItemType: {0}.", listType.ToString()));
			}

			var abstractNumTemplate = listTemplate.Descendants().Single(d => d.Name.LocalName == "abstractNum");
			abstractNumTemplate.SetAttributeValue(DocX.w + "abstractNumId", abstractNumId);

			//Fixing an issue where numbering would continue from previous numbered lists. Setting startOverride assures that a numbered list starts on the provided number.
			//The override needs only be on level 0 as this will cascade to the rest of the list.
			var abstractNumXml = GetAbstractNumXml(abstractNumId, numId, startNumber, continueNumbering);

			var abstractNumNode = Document.numbering.Root.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "abstractNum");
			var numXml = Document.numbering.Root.Descendants().LastOrDefault(xElement => xElement.Name.LocalName == "num");

			if (abstractNumNode == null || numXml == null)
			{
				Document.numbering.Root.Add(abstractNumTemplate);
				Document.numbering.Root.Add(abstractNumXml);
			}
			else
			{
				abstractNumNode.AddAfterSelf(abstractNumTemplate);
				numXml.AddAfterSelf(
					abstractNumXml
				);
			}

			NumId = numId;
		}

		private XElement GetAbstractNumXml(int abstractNumId, int numId, int? startNumber, bool continueNumbering)
		{
			//Fixing an issue where numbering would continue from previous numbered lists. Setting startOverride assures that a numbered list starts on the provided number.
			//The override needs only be on level 0 as this will cascade to the rest of the list.
			var startOverride = new XElement(XName.Get("startOverride", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", startNumber ?? 1));
			var lvlOverride = new XElement(XName.Get("lvlOverride", DocX.w.NamespaceName), new XAttribute(DocX.w + "ilvl", 0), startOverride);
			var abstractNumIdElement = new XElement(XName.Get("abstractNumId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", abstractNumId));
			return continueNumbering
				? new XElement(XName.Get("num", DocX.w.NamespaceName), new XAttribute(DocX.w + "numId", numId), abstractNumIdElement)
				: new XElement(XName.Get("num", DocX.w.NamespaceName), new XAttribute(DocX.w + "numId", numId), abstractNumIdElement, lvlOverride);
		}

		/// <summary>
		/// Method to determine the last numId for a list element. 
		/// Also useful for determining the next numId to use for inserting a new list element into the document.
		/// </summary>
		/// <returns>
		/// 0 if there are no elements in the list already.
		/// Increment the return for the next valid value of a new list element.
		/// </returns>
		private int GetMaxNumId()
		{
			const int defaultValue = 0;
			if (Document.numbering == null)
				return defaultValue;

			var numlist = Document.numbering.Descendants().Where(d => d.Name.LocalName == "num").ToList();
			if (numlist.Any())
				return numlist.Attributes(DocX.w + "numId").Max(e => int.Parse(e.Value));
			return defaultValue;
		}

		/// <summary>
		/// Method to determine the last abstractNumId for a list element.
		/// Also useful for determining the next abstractNumId to use for inserting a new list element into the document.
		/// </summary>
		/// <returns>
		/// -1 if there are no elements in the list already.
		/// Increment the return for the next valid value of a new list element.
		/// </returns>
		private int GetMaxAbstractNumId()
		{
			const int defaultValue = -1;

			if (Document.numbering == null)
				return defaultValue;

			var numlist = Document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum").ToList();
			if (numlist.Any())
			{
				var maxAbstractNumId = numlist.Attributes(DocX.w + "abstractNumId").Max(e => int.Parse(e.Value));
				return maxAbstractNumId;
			}
			return defaultValue;
		}

		/// <summary>
		/// Get the abstractNum definition for the given numId
		/// </summary>
		/// <param name="numId">The numId on the pPr element</param>
		/// <returns>XElement representing the requested abstractNum</returns>
		internal XElement GetAbstractNum(int numId)
		{
			var num = Document.numbering.Descendants().First(d => d.Name.LocalName == "num" && d.GetAttribute(DocX.w + "numId").Equals(numId.ToString()));
			var abstractNumId = num.Descendants().First(d => d.Name.LocalName == "abstractNumId");
			return Document.numbering.Descendants().First(d => d.Name.LocalName == "abstractNum" && d.GetAttribute("abstractNumId").Equals(abstractNumId.Value));
		}

		private void ValidateDocXNumberingPartExists()
		{
			var numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);

			// If the internal document contains no /word/numbering.xml create one.
			if (!Document.package.PartExists(numberingUri))
				Document.numbering = HelperFunctions.AddDefaultNumberingXml(Document.package);
		}
	}

	/// <summary>See <a href="https://support.microsoft.com/en-gb/kb/951731" /> for explanation</summary>
	public class PackagePartStream : Stream
	{

		private static readonly Mutex Mutex = new Mutex(false);

		private readonly Stream stream;

		/// <summary></summary>
		/// <param name="stream"></param>
		public PackagePartStream(Stream stream)
		{
			this.stream = stream;
		}

		/// <summary></summary>
		public override bool CanRead
		{
			get
			{
				return stream.CanRead;
			}
		}

		/// <summary></summary>
		public override bool CanSeek
		{
			get
			{
				return this.stream.CanSeek;
			}
		}

		/// <summary></summary>
		public override bool CanWrite
		{
			get
			{
				return this.stream.CanWrite;
			}
		}

		/// <summary></summary>
		public override long Length
		{
			get { return this.stream.Length; }
		}

		/// <summary></summary>
		public override long Position
		{
			get { return this.stream.Position; }
			set { this.stream.Position = value; }
		}

		/// <summary></summary>
		/// <param name="offset"></param>
		/// <param name="origin"></param>
		/// <returns></returns>
		public override long Seek(long offset, SeekOrigin origin)
		{
			return this.stream.Seek(offset, origin);
		}

		/// <summary></summary>
		/// <param name="value"></param>
		public override void SetLength(long value)
		{
			this.stream.SetLength(value);
		}

		/// <summary></summary>
		/// <param name="buffer"></param>
		/// <param name="offset"></param>
		/// <param name="count"></param>
		/// <returns></returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			return this.stream.Read(buffer, offset, count);
		}

		/// <summary></summary>
		/// <param name="buffer"></param>
		/// <param name="offset"></param>
		/// <param name="count"></param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			Mutex.WaitOne(Timeout.Infinite, false);
			this.stream.Write(buffer, offset, count);
			Mutex.ReleaseMutex();
		}

		/// <summary></summary>
		public override void Flush()
		{
			Mutex.WaitOne(Timeout.Infinite, false);
			this.stream.Flush();
			Mutex.ReleaseMutex();
		}

		/// <summary></summary>
		public override void Close()
		{
			this.stream.Close();
		}

		/// <summary></summary>
		/// <param name="disposing"></param>
		protected override void Dispose(bool disposing)
		{
			this.stream.Dispose();
		}

	}

	/// <summary></summary>
	public class PageLayout : DocXElement
	{

		internal PageLayout(DocX document, XElement xml) : base(document, xml)
		{
		}

		/// <summary></summary>
		public Orientation Orientation
		{
			get
			{
				/*
                 * Get the pgSz (page size) element for this Section,
                 * null will be return if no such element exists.
                 */
				XElement pgSz = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));

				if (pgSz == null)
					return Orientation.Portrait;

				// Get the attribute of the pgSz element.
				XAttribute val = pgSz.Attribute(XName.Get("orient", DocX.w.NamespaceName));

				// If val is null, this cell contains no information.
				if (val == null)
					return Orientation.Portrait;

				if (val.Value.Equals("Landscape", StringComparison.CurrentCultureIgnoreCase))
					return Orientation.Landscape;
				else
					return Orientation.Portrait;
			}

			set
			{
				// Check if already correct value.
				if (Orientation == value)
					return;

				/*
                 * Get the pgSz (page size) element for this Section,
                 * null will be return if no such element exists.
                 */
				XElement pgSz = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));

				if (pgSz == null)
				{
					Xml.SetElementValue(XName.Get("pgSz", DocX.w.NamespaceName), string.Empty);
					pgSz = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));
				}

				pgSz.SetAttributeValue(XName.Get("orient", DocX.w.NamespaceName), value.ToString().ToLower());

				if (value == Novacode.Orientation.Landscape)
				{
					pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "16838");
					pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "11906");
				}

				else if (value == Novacode.Orientation.Portrait)
				{
					pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "11906");
					pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "16838");
				}
			}
		}
	}

	/// <summary>Represents a document paragraph</summary>
	public class Paragraph : InsertBeforeOrAfter
	{

		// The Append family of functions use this List to apply style.
		internal List<XElement> runs;

		// This paragraphs text alignment
		private Alignment alignment;

		/// <summary></summary>
		public ContainerType ParentContainer;

		private XElement ParagraphNumberPropertiesBacker { get; set; }

		/// <summary>
		/// Fetch the paragraph number properties for a list element.
		/// </summary>
		public XElement ParagraphNumberProperties
		{
			get
			{
				return ParagraphNumberPropertiesBacker ?? (ParagraphNumberPropertiesBacker = GetParagraphNumberProperties());
			}
		}

		private XElement GetParagraphNumberProperties()
		{
			var numPrNode = Xml.Descendants().FirstOrDefault(el => el.Name.LocalName == "numPr");
			if (numPrNode != null)
			{
				var numIdNode = numPrNode.Descendants().First(numId => numId.Name.LocalName == "numId");
				var numIdAttribute = numIdNode.Attribute(DocX.w + "val");
				if (numIdAttribute != null && numIdAttribute.Value.Equals("0"))
				{
					return null;
				}
			}

			return numPrNode;
		}

		private bool? IsListItemBacker { get; set; }
		/// <summary>
		/// Determine if this paragraph is a list element.
		/// </summary>
		public bool IsListItem
		{
			get
			{
				IsListItemBacker = IsListItemBacker ?? (ParagraphNumberProperties != null);
				return (bool)IsListItemBacker;
			}
		}

		private int? IndentLevelBacker { get; set; }
		/// <summary>
		/// If this element is a list item, get the indentation level of the list item.
		/// </summary>
		public int? IndentLevel
		{
			get
			{
				if (!IsListItem)
				{
					return null;
				}
				return IndentLevelBacker ?? (IndentLevelBacker = int.Parse(ParagraphNumberProperties.Descendants().First(el => el.Name.LocalName == "ilvl").GetAttribute(DocX.w + "val")));
			}
		}

		/// <summary>
		/// Determine if the list element is a numbered list of bulleted list element
		/// </summary>
		public ListItemType ListItemType;

		internal int startIndex, endIndex;

		/// <summary>
		/// Returns a list of all Pictures in a Paragraph.
		/// </summary>
		/// <example>
		/// Returns a list of all Pictures in a Paragraph.
		/// <code>
		/// <![CDATA[
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///    // Get the first Paragraph in a document.
		///    Paragraph p = document.Paragraphs[0];
		/// 
		///    // Get all of the Pictures in this Paragraph.
		///    List<Picture> pictures = p.Pictures;
		///
		///    // Save this document.
		///    document.Save();
		/// }
		/// ]]>
		/// </code>
		/// </example>
		public List<Picture> Pictures
		{
			get
			{
				List<Picture> pictures =
				(
					from p in Xml.Descendants()
					where (p.Name.LocalName == "drawing")
					let id =
					(
						from e in p.Descendants()
						where e.Name.LocalName.Equals("blip")
						select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
					).SingleOrDefault()
					where id != null
					let img = new Image(Document, mainPart.GetRelationship(id))
					select new Picture(Document, p, img)
				).ToList();

				List<Picture> shapes =
				(
					from p in Xml.Descendants()
					where (p.Name.LocalName == "pict")
					let id =
					(
						from e in p.Descendants()
						where e.Name.LocalName.Equals("imagedata")
						select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
					).SingleOrDefault()
					where id != null
					let img = new Image(Document, mainPart.GetRelationship(id))
					select new Picture(Document, p, img)
				).ToList();

				foreach (Picture p in shapes)
					pictures.Add(p);


				return pictures;
			}
		}

		/// <summary>
		/// Returns a list of Hyperlinks in this Paragraph.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///    // Get the first Paragraph in this document.
		///    Paragraph p = document.Paragraphs[0];
		///    
		///    // Get all of the hyperlinks in this Paragraph.
		///    <![CDATA[ List<Hyperlink> ]]> hyperlinks = paragraph.Hyperlinks;
		///    
		///    // Change the first hyperlinks text and Uri
		///    Hyperlink h0 = hyperlinks[0];
		///    h0.Text = "DocX";
		///    h0.Uri = new Uri("http://docx.codeplex.com");
		///
		///    // Save this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public List<Hyperlink> Hyperlinks
		{
			get
			{
				List<Hyperlink> hyperlinks = new List<Hyperlink>();

				List<XElement> hyperlink_elements =
				(
					from h in Xml.Descendants()
					where (h.Name.LocalName == "hyperlink" || h.Name.LocalName == "instrText")
					select h
				).ToList();

				foreach (XElement he in hyperlink_elements)
				{
					if (he.Name.LocalName == "hyperlink")
					{
						try
						{
							Hyperlink h = new Hyperlink(Document, mainPart, he);
							h.mainPart = mainPart;
							hyperlinks.Add(h);
						}

						catch (Exception) { }
					}

					else
					{
						// Find the parent run, no matter how deeply nested we are.
						XElement e = he;
						while (e.Name.LocalName != "r")
							e = e.Parent;

						// Take every element until we reach w:fldCharType="end"
						List<XElement> hyperlink_runs = new List<XElement>();
						foreach (XElement r in e.ElementsAfterSelf(XName.Get("r", DocX.w.NamespaceName)))
						{
							// Add this run to the list.
							hyperlink_runs.Add(r);

							XElement fldChar = r.Descendants(XName.Get("fldChar", DocX.w.NamespaceName)).SingleOrDefault<XElement>();
							if (fldChar != null)
							{
								XAttribute fldCharType = fldChar.Attribute(XName.Get("fldCharType", DocX.w.NamespaceName));
								if (fldCharType != null && fldCharType.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase))
								{
									try
									{
										Hyperlink h = new Hyperlink(Document, he, hyperlink_runs);
										h.mainPart = mainPart;
										hyperlinks.Add(h);
									}

									catch (Exception) { }

									break;
								}
							}
						}
					}
				}

				return hyperlinks;
			}
		}

		///<summary>
		/// The style name of the paragraph.
		///</summary>
		public string StyleName
		{
			get
			{
				var element = this.GetOrCreate_pPr();
				var styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName));
				if (styleElement != null)
				{
					var attr = styleElement.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (attr != null && !string.IsNullOrEmpty(attr.Value))
					{
						return attr.Value;
					}
				}
				return "Normal";
			}
			set
			{
				if (string.IsNullOrEmpty(value))
				{
					value = "Normal";
				}
				var element = this.GetOrCreate_pPr();
				var styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName));
				if (styleElement == null)
				{
					element.Add(new XElement(XName.Get("pStyle", DocX.w.NamespaceName)));
					styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName));
				}
				styleElement.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value);
			}
		}

		// A collection of field type DocProperty.
		private List<DocProperty> docProperties;

		internal List<XElement> styles = new List<XElement>();

		/// <summary>
		/// Returns a list of field type DocProperty in this document.
		/// </summary>
		public List<DocProperty> DocumentProperties
		{
			get { return docProperties; }
		}

		internal Paragraph(DocX document, XElement xml, int startIndex, ContainerType parent = ContainerType.None)
			: base(document, xml)
		{
			ParentContainer = parent;
			this.startIndex = startIndex;
			this.endIndex = startIndex + GetElementTextLength(xml);

			RebuildDocProperties();

			// As per Unused code affecting performance (Wiki Link: [discussion:454191]) and coffeycathal suggestion no longer requeried
			//#region It's possible that a Paragraph may have pStyle references
			//// Check if this Paragraph references any pStyle elements.
			//var stylesElements = xml.Descendants(XName.Get("pStyle", DocX.w.NamespaceName));

			//// If one or more pStyles are referenced.
			//if (stylesElements.Count() > 0)
			//{
			//    Uri style_package_uri = new Uri("/word/styles.xml", UriKind.Relative);
			//    PackagePart styles_document = document.package.GetPart(style_package_uri);

			//    using (TextReader tr = new StreamReader(styles_document.GetStream()))
			//    {
			//        XDocument style_document = XDocument.Load(tr);
			//        XElement styles_element = style_document.Element(XName.Get("styles", DocX.w.NamespaceName));

			//        var styles_element_ids = stylesElements.Select(e => e.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value);

			//        //foreach(string id in styles_element_ids)
			//        //{
			//        //    var style = 
			//        //    (
			//        //        from d in styles_element.Descendants()
			//        //        let styleId = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
			//        //        let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
			//        //        where type != null && type.Value == "paragraph" && styleId != null && styleId.Value == id
			//        //        select d
			//        //    ).First();

			//        //    styles.Add(style);
			//        //} 
			//    }
			//}
			//#endregion

			this.runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
		}

		/// <summary>
		/// Insert a new Table before this Paragraph, this Table can be from this document or another document.
		/// </summary>
		/// <param name="t">The Table t to be inserted.</param>
		/// <returns>A new Table inserted before this Paragraph.</returns>
		/// <example>
		/// Insert a new Table before this Paragraph.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Paragraph in document b.
		///     Paragraph p2 = documentB.Paragraphs[0];
		///
		///     // Insert the Table from document a before this Paragraph.
		///     Table newTable = p2.InsertTableBeforeSelf(t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableBeforeSelf(Table t)
		{
			t = base.InsertTableBeforeSelf(t);
			t.mainPart = mainPart;
			return t;
		}

		private Direction direction;
		/// <summary>
		/// Gets or Sets the Direction of content in this Paragraph.
		/// <example>
		/// Create a Paragraph with content that flows right to left. Default is left to right.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Create a new Paragraph with the text "Hello World".
		///     Paragraph p = document.InsertParagraph("Hello World.");
		/// 
		///     // Make this Paragraph flow right to left. Default is left to right.
		///     p.Direction = Direction.RightToLeft;
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// </summary>
		public Direction Direction
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement bidi = pPr.Element(XName.Get("bidi", DocX.w.NamespaceName));

				if (bidi == null)
					return Direction.LeftToRight;

				else
					return Direction.RightToLeft;
			}

			set
			{
				direction = value;

				XElement pPr = GetOrCreate_pPr();
				XElement bidi = pPr.Element(XName.Get("bidi", DocX.w.NamespaceName));

				if (direction == Direction.RightToLeft)
				{
					if (bidi == null)
						pPr.Add(new XElement(XName.Get("bidi", DocX.w.NamespaceName)));
				}

				else
				{
					if (bidi != null)
						bidi.Remove();
				}
			}
		}

		/// <summary></summary>
		public bool IsKeepWithNext
		{

			get
			{
				var pPr = GetOrCreate_pPr();
				var keepWithNextE = pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName));
				if (keepWithNextE == null)
				{
					return false;
				}
				return true;
			}
		}
		/// <summary>
		/// This paragraph will be kept on the same page as the next paragraph
		/// </summary>
		/// <example>
		/// Create a Paragraph that will stay on the same page as the paragraph that comes next
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// 
		/// {
		///     // Create a new Paragraph with the text "Hello World".
		///     Paragraph p = document.InsertParagraph("Hello World.");
		///     p.KeepWithNext();
		///     document.InsertParagraph("Previous paragraph will appear on the same page as this paragraph");
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <param name="keepWithNext"></param>
		/// <returns></returns>

		public Paragraph KeepWithNext(bool keepWithNext = true)
		{
			var pPr = GetOrCreate_pPr();
			var keepWithNextE = pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName));
			if (keepWithNextE == null && keepWithNext)
			{
				pPr.Add(new XElement(XName.Get("keepNext", DocX.w.NamespaceName)));
			}
			if (!keepWithNext && keepWithNextE != null)
			{
				keepWithNextE.Remove();
			}
			return this;

		}
		/// <summary>
		/// Keep all lines in this paragraph together on a page
		/// </summary>
		/// <example>
		/// Create a Paragraph whose lines will stay together on a single page
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Create a new Paragraph with the text "Hello World".
		///     Paragraph p = document.InsertParagraph("All lines of this paragraph will appear on the same page...\nLine 2\nLine 3\nLine 4\nLine 5\nLine 6...");
		///     p.KeepLinesTogether();
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <param name="keepTogether"></param>
		/// <returns></returns>
		public Paragraph KeepLinesTogether(bool keepTogether = true)
		{
			var pPr = GetOrCreate_pPr();
			var keepLinesE = pPr.Element(XName.Get("keepLines", DocX.w.NamespaceName));
			if (keepLinesE == null && keepTogether)
			{
				pPr.Add(new XElement(XName.Get("keepLines", DocX.w.NamespaceName)));
			}
			if (!keepTogether && keepLinesE != null)
			{
				keepLinesE.Remove();
			}
			return this;
		}
		/// <summary>
		/// If the pPr element doesent exist it is created, either way it is returned by this function.
		/// </summary>
		/// <returns>The pPr element for this Paragraph.</returns>
		internal XElement GetOrCreate_pPr()
		{
			// Get the element.
			XElement pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));

			// If it dosen't exist, create it.
			if (pPr == null)
			{
				Xml.AddFirst(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
				pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
			}

			// Return the pPr element for this Paragraph.
			return pPr;
		}

		/// <summary>
		/// If the ind element doesent exist it is created, either way it is returned by this function.
		/// </summary>
		/// <returns>The ind element for this Paragraphs pPr.</returns>
		internal XElement GetOrCreate_pPr_ind()
		{
			// Get the element.
			XElement pPr = GetOrCreate_pPr();
			XElement ind = pPr.Element(XName.Get("ind", DocX.w.NamespaceName));

			// If it dosen't exist, create it.
			if (ind == null)
			{
				pPr.Add(new XElement(XName.Get("ind", DocX.w.NamespaceName)));
				ind = pPr.Element(XName.Get("ind", DocX.w.NamespaceName));
			}

			// Return the pPr element for this Paragraph.
			return ind;
		}

		private float indentationFirstLine;
		/// <summary>
		/// Get or set the indentation of the first line of this Paragraph.
		/// </summary>
		/// <example>
		/// Indent only the first line of a Paragraph.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Create a new Paragraph.
		///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		/// 
		///     // Indent only the first line of the Paragraph.
		///     p.IndentationFirstLine = 2.0f;
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public float IndentationFirstLine
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement ind = GetOrCreate_pPr_ind();
				XAttribute firstLine = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName));

				if (firstLine != null)
					return float.Parse(firstLine.Value);

				return 0.0f;
			}

			set
			{
				if (IndentationFirstLine != value)
				{
					indentationFirstLine = value;

					XElement pPr = GetOrCreate_pPr();
					XElement ind = GetOrCreate_pPr_ind();

					// Paragraph can either be firstLine or hanging (Remove hanging).
					XAttribute hanging = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName));
					if (hanging != null)
						hanging.Remove();

					string indentation = ((indentationFirstLine / 0.1) * 57).ToString();
					XAttribute firstLine = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName));
					if (firstLine != null)
						firstLine.Value = indentation;
					else
						ind.Add(new XAttribute(XName.Get("firstLine", DocX.w.NamespaceName), indentation));
				}
			}
		}

		private float indentationHanging;
		/// <summary>
		/// Get or set the indentation of all but the first line of this Paragraph.
		/// </summary>
		/// <example>
		/// Indent all but the first line of a Paragraph.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Create a new Paragraph.
		///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		/// 
		///     // Indent all but the first line of the Paragraph.
		///     p.IndentationHanging = 1.0f;
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public float IndentationHanging
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement ind = GetOrCreate_pPr_ind();
				XAttribute hanging = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName));

				if (hanging != null)
					return float.Parse(hanging.Value) / (57 * 10);

				return 0.0f;
			}

			set
			{
				if (IndentationHanging != value)
				{
					indentationHanging = value;

					XElement pPr = GetOrCreate_pPr();
					XElement ind = GetOrCreate_pPr_ind();

					// Paragraph can either be firstLine or hanging (Remove firstLine).
					XAttribute firstLine = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName));
					if (firstLine != null)
						firstLine.Remove();

					string indentation = ((indentationHanging / 0.1) * 57).ToString();
					XAttribute hanging = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName));
					if (hanging != null)
						hanging.Value = indentation;
					else
						ind.Add(new XAttribute(XName.Get("hanging", DocX.w.NamespaceName), indentation));
				}
			}
		}

		private float indentationBefore;
		/// <summary>
		/// Set the before indentation in cm for this Paragraph.
		/// </summary>
		/// <example>
		/// // Indent an entire Paragraph from the left.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///    // Create a new Paragraph.
		///    Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		///
		///    // Indent this entire Paragraph from the left.
		///    p.IndentationBefore = 2.0f;
		///    
		///    // Save all changes made to this document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public float IndentationBefore
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement ind = GetOrCreate_pPr_ind();

				XAttribute left = ind.Attribute(XName.Get("left", DocX.w.NamespaceName));
				if (left != null)
					return float.Parse(left.Value) / (57 * 10);

				return 0.0f;
			}

			set
			{
				if (IndentationBefore != value)
				{
					indentationBefore = value;

					XElement pPr = GetOrCreate_pPr();
					XElement ind = GetOrCreate_pPr_ind();

					string indentation = ((indentationBefore / 0.1) * 57).ToString();

					XAttribute left = ind.Attribute(XName.Get("left", DocX.w.NamespaceName));
					if (left != null)
						left.Value = indentation;
					else
						ind.Add(new XAttribute(XName.Get("left", DocX.w.NamespaceName), indentation));
				}
			}
		}

		private float indentationAfter = 0.0f;
		/// <summary>
		/// Set the after indentation in cm for this Paragraph.
		/// </summary>
		/// <example>
		/// // Indent an entire Paragraph from the right.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Create a new Paragraph.
		///     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		/// 
		///     // Make the content of this Paragraph flow right to left.
		///     p.Direction = Direction.RightToLeft;
		/// 
		///     // Indent this entire Paragraph from the right.
		///     p.IndentationAfter = 2.0f;
		///     
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public float IndentationAfter
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement ind = GetOrCreate_pPr_ind();

				XAttribute right = ind.Attribute(XName.Get("right", DocX.w.NamespaceName));
				if (right != null)
					return float.Parse(right.Value);

				return 0.0f;
			}

			set
			{
				if (IndentationAfter != value)
				{
					indentationAfter = value;

					XElement pPr = GetOrCreate_pPr();
					XElement ind = GetOrCreate_pPr_ind();

					string indentation = ((indentationAfter / 0.1) * 57).ToString();

					XAttribute right = ind.Attribute(XName.Get("right", DocX.w.NamespaceName));
					if (right != null)
						right.Value = indentation;
					else
						ind.Add(new XAttribute(XName.Get("right", DocX.w.NamespaceName), indentation));
				}
			}
		}

		/// <summary>
		/// Insert a new Table into this document before this Paragraph.
		/// </summary>
		/// <param name="rowCount">The number of rows this Table should have.</param>
		/// <param name="columnCount">The number of columns this Table should have.</param>
		/// <returns>A new Table inserted before this Paragraph.</returns>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     //Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("Hello World", false);
		///
		///     // Insert a new Table before this Paragraph.
		///     Table newTable = p.InsertTableBeforeSelf(2, 2);
		///     newTable.Design = TableDesign.LightShadingAccent2;
		///     newTable.Alignment = Alignment.center;
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableBeforeSelf(int rowCount, int columnCount)
		{
			return base.InsertTableBeforeSelf(rowCount, columnCount);
		}

		/// <summary>
		/// Insert a new Table after this Paragraph.
		/// </summary>
		/// <param name="t">The Table t to be inserted.</param>
		/// <returns>A new Table inserted after this Paragraph.</returns>
		/// <example>
		/// Insert a new Table after this Paragraph.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Paragraph in document b.
		///     Paragraph p2 = documentB.Paragraphs[0];
		///
		///     // Insert the Table from document a after this Paragraph.
		///     Table newTable = p2.InsertTableAfterSelf(t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableAfterSelf(Table t)
		{
			t = base.InsertTableAfterSelf(t);
			t.mainPart = mainPart;
			return t;
		}

		/// <summary>
		/// Insert a new Table into this document after this Paragraph.
		/// </summary>
		/// <param name="rowCount">The number of rows this Table should have.</param>
		/// <param name="columnCount">The number of columns this Table should have.</param>
		/// <returns>A new Table inserted after this Paragraph.</returns>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     //Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("Hello World", false);
		///
		///     // Insert a new Table after this Paragraph.
		///     Table newTable = p.InsertTableAfterSelf(2, 2);
		///     newTable.Design = TableDesign.LightShadingAccent2;
		///     newTable.Alignment = Alignment.center;
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableAfterSelf(int rowCount, int columnCount)
		{
			return base.InsertTableAfterSelf(rowCount, columnCount);
		}

		/// <summary>
		/// Insert a Paragraph before this Paragraph, this Paragraph may have come from the same or another document.
		/// </summary>
		/// <param name="p">The Paragraph to insert.</param>
		/// <returns>The Paragraph now associated with this document.</returns>
		/// <example>
		/// Take a Paragraph from document a, and insert it into document b before this Paragraph.
		/// <code>
		/// // Place holder for a Paragraph.
		/// Paragraph p;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first paragraph from this document.
		///     p = documentA.Paragraphs[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Paragraph in document b.
		///     Paragraph p2 = documentB.Paragraphs[0];
		///
		///     // Insert the Paragraph from document a before this Paragraph.
		///     Paragraph newParagraph = p2.InsertParagraphBeforeSelf(p);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code> 
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
		{
			Paragraph p2 = base.InsertParagraphBeforeSelf(p);
			p2.PackagePart = mainPart;
			return p2;
		}

		/// <summary>
		/// Insert a new Paragraph before this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <returns>A new Paragraph inserted before this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph before the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.");
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text)
		{
			Paragraph p = base.InsertParagraphBeforeSelf(text);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>
		/// Insert a new Paragraph before this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <returns>A new Paragraph inserted before this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph before the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraphBeforeSelf(text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>
		/// Insert a new Paragraph before this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <param name="formatting">The formatting to apply to this insertion.</param>
		/// <returns>A new Paragraph inserted before this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph before the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     Formatting boldFormatting = new Formatting();
		///     boldFormatting.Bold = true;
		///
		///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false, boldFormatting);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>
		/// Insert a page break before a Paragraph.
		/// </summary>
		/// <example>
		/// Insert 2 Paragraphs into a document with a page break between them.
		/// <code>
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Insert a new Paragraph.
		///    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
		///       
		///    // Insert a new Paragraph.
		///    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
		///    
		///    // Insert a page break before Paragraph two.
		///    p2.InsertPageBreakBeforeSelf();
		///    
		///    // Save this document.
		///    document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override void InsertPageBreakBeforeSelf()
		{
			base.InsertPageBreakBeforeSelf();
		}

		/// <summary>
		/// Insert a page break after a Paragraph.
		/// </summary>
		/// <example>
		/// Insert 2 Paragraphs into a document with a page break between them.
		/// <code>
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Insert a new Paragraph.
		///    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
		///       
		///    // Insert a page break after this Paragraph.
		///    p1.InsertPageBreakAfterSelf();
		///       
		///    // Insert a new Paragraph.
		///    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
		///
		///    // Save this document.
		///    document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override void InsertPageBreakAfterSelf()
		{
			base.InsertPageBreakAfterSelf();
		}

		/// <summary></summary>
		/// <param name="index"></param>
		/// <param name="h"></param>
		/// <returns></returns>
		[Obsolete("Instead use: InsertHyperlink(Hyperlink h, int index)")]
		public Paragraph InsertHyperlink(int index, Hyperlink h) { return InsertHyperlink(h, index); }

		/// <summary>
		/// This function inserts a hyperlink into a Paragraph at a specified character index.
		/// </summary>
		/// <param name="index">The index to insert at.</param>
		/// <param name="h">The hyperlink to insert.</param>
		/// <returns>The Paragraph with the Hyperlink inserted at the specified index.</returns>
		/// <!-- 
		/// This function was added by Brian Campbell aka chickendelicious on Jun 16 2010
		/// Thank you Brian.
		/// -->
		public Paragraph InsertHyperlink(Hyperlink h, int index = 0)
		{
			// Convert the path of this mainPart to its equilivant rels file path.
			string path = mainPart.Uri.OriginalString.Replace("/word/", "");
			Uri rels_path = new Uri(String.Format("/word/_rels/{0}.rels", path), UriKind.Relative);

			// Check to see if the rels file exists and create it if not.
			if (!Document.package.PartExists(rels_path))
				HelperFunctions.CreateRelsPackagePart(Document, rels_path);

			// Check to see if a rel for this Picture exists, create it if not.
			var Id = GetOrGenerateRel(h);

			XElement h_xml;
			if (index == 0)
			{
				// Add this hyperlink as the last element.
				Xml.AddFirst(h.Xml);

				// Extract the picture back out of the DOM.
				h_xml = (XElement)Xml.FirstNode;
			}

			else
			{
				// Get the first run effected by this Insert
				Run run = GetFirstRunEffectedByEdit(index);

				if (run == null)
				{
					// Add this hyperlink as the last element.
					Xml.Add(h.Xml);

					// Extract the picture back out of the DOM.
					h_xml = (XElement)Xml.LastNode;
				}

				else
				{
					// Split this run at the point you want to insert
					XElement[] splitRun = Run.SplitRun(run, index);

					// Replace the origional run.
					run.Xml.ReplaceWith
					(
						splitRun[0],
						h.Xml,
						splitRun[1]
					);

					// Get the first run effected by this Insert
					run = GetFirstRunEffectedByEdit(index);

					// The picture has to be the next element, extract it back out of the DOM.
					h_xml = (XElement)run.Xml.NextNode;
				}

				h_xml.SetAttributeValue(DocX.r + "id", Id);
			}

			return this;
		}

		/// <summary>
		/// Remove the Hyperlink at the provided index. The first hyperlink is at index 0.
		/// Using a negative index or an index greater than the index of the last hyperlink will cause an ArgumentOutOfRangeException() to be thrown.
		/// </summary>
		/// <param name="index">The index of the hyperlink to be removed.</param>
		/// <example>
		/// <code>
		/// // Crete a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///     // Add a Hyperlink into this document.
		///     Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
		///
		///     // Insert a new Paragraph into the document.
		///     Paragraph p1 = document.InsertParagraph("AC");
		///     
		///     // Insert the hyperlink into this Paragraph.
		///     p1.InsertHyperlink(1, h);
		///     Assert.IsTrue(p1.Text == "AlinkC"); // Make sure the hyperlink was inserted correctly;
		///     
		///     // Remove the hyperlink
		///     p1.RemoveHyperlink(0);
		///     Assert.IsTrue(p1.Text == "AC"); // Make sure the hyperlink was removed correctly;
		/// }
		/// </code>
		/// </example>
		public void RemoveHyperlink(int index)
		{
			// Dosen't make sense to remove a Hyperlink at a negative index.
			if (index < 0)
				throw new ArgumentOutOfRangeException();

			// Need somewhere to store the count.
			int count = 0;
			bool found = false;
			RemoveHyperlinkRecursive(Xml, index, ref count, ref found);

			// If !found then the user tried to remove a hyperlink at an index greater than the last. 
			if (!found)
				throw new ArgumentOutOfRangeException();
		}

		internal void RemoveHyperlinkRecursive(XElement xml, int index, ref int count, ref bool found)
		{
			if (xml.Name.LocalName.Equals("hyperlink", StringComparison.CurrentCultureIgnoreCase))
			{
				// This is the hyperlink to be removed.
				if (count == index)
				{
					found = true;
					xml.Remove();
				}

				else
					count++;
			}

			if (xml.HasElements)
				foreach (XElement e in xml.Elements())
					if (!found)
						RemoveHyperlinkRecursive(e, index, ref count, ref found);
		}

		/// <summary>
		/// Insert a Paragraph after this Paragraph, this Paragraph may have come from the same or another document.
		/// </summary>
		/// <param name="p">The Paragraph to insert.</param>
		/// <returns>The Paragraph now associated with this document.</returns>
		/// <example>
		/// Take a Paragraph from document a, and insert it into document b after this Paragraph.
		/// <code>
		/// // Place holder for a Paragraph.
		/// Paragraph p;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first paragraph from this document.
		///     p = documentA.Paragraphs[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Paragraph in document b.
		///     Paragraph p2 = documentB.Paragraphs[0];
		///
		///     // Insert the Paragraph from document a after this Paragraph.
		///     Paragraph newParagraph = p2.InsertParagraphAfterSelf(p);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code> 
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(Paragraph p)
		{
			Paragraph p2 = base.InsertParagraphAfterSelf(p);
			p2.PackagePart = mainPart;
			return p2;
		}

		/// <summary>
		/// Insert a new Paragraph after this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <param name="formatting">The formatting to apply to this insertion.</param>
		/// <returns>A new Paragraph inserted after this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph after the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     Formatting boldFormatting = new Formatting();
		///     boldFormatting.Bold = true;
		///
		///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false, boldFormatting);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph p = base.InsertParagraphAfterSelf(text, trackChanges, formatting);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>
		/// Insert a new Paragraph after this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <returns>A new Paragraph inserted after this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph after the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
		{
			Paragraph p = base.InsertParagraphAfterSelf(text, trackChanges);
			p.PackagePart = mainPart;
			return p;
		}

		/// <summary>
		/// Insert a new Paragraph after this Paragraph.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <returns>A new Paragraph inserted after this Paragraph.</returns>
		/// <example>
		/// Insert a new paragraph after the first Paragraph in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		///
		///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.");
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text)
		{
			Paragraph p = base.InsertParagraphAfterSelf(text);
			p.PackagePart = mainPart;
			return p;
		}

		private void RebuildDocProperties()
		{
			docProperties =
			(
				from xml in Xml.Descendants(XName.Get("fldSimple", DocX.w.NamespaceName))
				select new DocProperty(Document, xml)
			).ToList();
		}

		/// <summary>
		/// Gets or set this Paragraphs text alignment.
		/// </summary>
		public Alignment Alignment
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

				if (jc != null)
				{
					XAttribute a = jc.Attribute(XName.Get("val", DocX.w.NamespaceName));

					switch (a.Value.ToLower())
					{
						case "left": return Novacode.Alignment.left;
						case "right": return Novacode.Alignment.right;
						case "center": return Novacode.Alignment.center;
						case "both": return Novacode.Alignment.both;
					}
				}

				return Novacode.Alignment.left;
			}

			set
			{
				alignment = value;

				XElement pPr = GetOrCreate_pPr();
				XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

				if (alignment != Novacode.Alignment.left)
				{
					if (jc == null)
						pPr.Add(new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignment.ToString())));
					else
						jc.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value = alignment.ToString();
				}

				else
				{
					if (jc != null)
						jc.Remove();
				}
			}
		}

		/// <summary>
		/// Remove this Paragraph from the document.
		/// </summary>
		/// <param name="trackChanges">Should this remove be tracked as a change?</param>
		/// <example>
		/// Remove a Paragraph from a document and track it as a change.
		/// <code>
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Create and Insert a new Paragraph into this document.
		///     Paragraph p = document.InsertParagraph("Hello", false);
		///
		///     // Remove the Paragraph and track this as a change.
		///     p.Remove(true);
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void Remove(bool trackChanges)
		{
			if (trackChanges)
			{
				DateTime now = DateTime.Now.ToUniversalTime();

				List<XElement> elements = Xml.Elements().ToList();
				List<XElement> temp = new List<XElement>();
				for (int i = 0; i < elements.Count(); i++)
				{
					XElement e = elements[i];

					if (e.Name.LocalName != "del")
					{
						temp.Add(e);
						e.Remove();
					}

					else
					{
						if (temp.Count() > 0)
						{
							e.AddBeforeSelf(CreateEdit(EditType.del, now, temp.Elements()));
							temp.Clear();
						}
					}
				}

				if (temp.Count() > 0)
					Xml.Add(CreateEdit(EditType.del, now, temp));
			}

			else
			{
				// If this is the only Paragraph in the Cell then we cannot remove it.
				if (Xml.Parent.Name.LocalName == "tc" && Xml.Parent.Elements(XName.Get("p", DocX.w.NamespaceName)).Count() == 1)
					Xml.Value = string.Empty;

				else
				{
					// Remove this paragraph from the document
					Xml.Remove();
					Xml = null;
				}
			}
		}

		/// <summary>
		/// Gets the text value of this Paragraph.
		/// </summary>
		public string Text
		{
			// Returns the underlying XElement's Value property.
			get
			{
				return HelperFunctions.GetText(Xml);
			}
		}

		/// <summary>
		/// Gets the formatted text value of this Paragraph.
		/// </summary>
		public List<FormattedText> MagicText
		{
			// Returns the underlying XElement's Value property.
			get
			{
				try
				{
					return HelperFunctions.GetFormattedText(Xml);

				}
				catch (Exception)
				{
					return null;
				}

			}
		}

		//public Picture InsertPicture(Picture picture)
		//{
		//    Picture newPicture = picture;
		//    newPicture.i = new XElement(picture.i);

		//    xml.Add(newPicture.i);
		//    pictures.Add(newPicture);
		//    return newPicture;  
		//}

		// <summary>
		// Insert a Picture at the end of this paragraph.
		// </summary>
		// <param name="description">A string to describe this Picture.</param>
		// <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
		// <param name="name">The name of this image.</param>
		// <returns>A Picture.</returns>
		// <example>
		// <code>
		// // Create a document using a relative filename.
		// using (DocX document = DocX.Create(@"Test.docx"))
		// {
		//     // Add a new Paragraph to this document.
		//     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
		//
		//     // Add an Image to this document.
		//     Novacode.Image img = document.AddImage(@"Image.jpg");
		//
		//     // Insert pic at the end of Paragraph p.
		//     Picture pic = p.InsertPicture(img.Id, "Photo 31415", "A pie I baked.");
		//
		//     // Rotate the Picture clockwise by 30 degrees. 
		//     pic.Rotation = 30;
		//
		//     // Resize the Picture.
		//     pic.Width = 400;
		//     pic.Height = 300;
		//
		//     // Set the shape of this Picture to be a cube.
		//     pic.SetPictureShape(BasicShapes.cube);
		//
		//     // Flip the Picture Horizontally.
		//     pic.FlipHorizontal = true;
		//
		//     // Save all changes made to this document.
		//     document.Save();
		// }// Release this document from memory.
		// </code>
		// </example>
		// Removed to simplify the API.
		//public Picture InsertPicture(string imageID, string name, string description)
		//{
		//    Picture p = CreatePicture(Document, imageID, name, description);
		//    Xml.Add(p.Xml);
		//    return p;
		//}

		// Removed because it confusses the API.
		//public Picture InsertPicture(string imageID)
		//{
		//    return InsertPicture(imageID, string.Empty, string.Empty);
		//}

		//public Picture InsertPicture(int index, Picture picture)
		//{
		//    Picture p = picture;
		//    p.i = new XElement(picture.i);

		//    Run run = GetFirstRunEffectedByEdit(index);

		//    if (run == null)
		//        xml.Add(p.i);
		//    else
		//    {
		//        // Split this run at the point you want to insert
		//        XElement[] splitRun = Run.SplitRun(run, index);

		//        // Replace the origional run
		//        run.Xml.ReplaceWith
		//        (
		//            splitRun[0],
		//            p.i,
		//            splitRun[1]
		//        );
		//    }

		//    // Rebuild the run lookup for this paragraph
		//    runLookup.Clear();
		//    BuildRunLookup(xml);
		//    DocX.RenumberIDs(document);
		//    return p;
		//}

		// <summary>
		// Insert a Picture into this Paragraph at a specified index.
		// </summary>
		// <param name="description">A string to describe this Picture.</param>
		// <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
		// <param name="name">The name of this image.</param>
		// <param name="index">The index to insert this Picture at.</param>
		// <returns>A Picture.</returns>
		// <example>
		// <code>
		// // Create a document using a relative filename.
		// using (DocX document = DocX.Create(@"Test.docx"))
		// {
		//     // Add a new Paragraph to this document.
		//     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
		//
		//     // Add an Image to this document.
		//     Novacode.Image img = document.AddImage(@"Image.jpg");
		//
		//     // Insert pic at the start of Paragraph p.
		//     Picture pic = p.InsertPicture(0, img.Id, "Photo 31415", "A pie I baked.");
		//
		//     // Rotate the Picture clockwise by 30 degrees. 
		//     pic.Rotation = 30;
		//
		//     // Resize the Picture.
		//     pic.Width = 400;
		//     pic.Height = 300;
		//
		//     // Set the shape of this Picture to be a cube.
		//     pic.SetPictureShape(BasicShapes.cube);
		//
		//     // Flip the Picture Horizontally.
		//     pic.FlipHorizontal = true;
		//
		//     // Save all changes made to this document.
		//     document.Save();
		// }// Release this document from memory.
		// </code>
		// </example>
		// Removed to simplify API.
		//public Picture InsertPicture(int index, string imageID, string name, string description)
		//{
		//    Picture picture = CreatePicture(Document, imageID, name, description);

		//    Run run = GetFirstRunEffectedByEdit(index);

		//    if (run == null)
		//        Xml.Add(picture.Xml);
		//    else
		//    {
		//        // Split this run at the point you want to insert
		//        XElement[] splitRun = Run.SplitRun(run, index);

		//        // Replace the origional run
		//        run.Xml.ReplaceWith
		//        (
		//            splitRun[0],
		//            picture.Xml,
		//            splitRun[1]
		//        );
		//    }

		//    HelperFunctions.RenumberIDs(Document);
		//    return picture;
		//}

		/// <summary>
		/// Create a new Picture.
		/// </summary>
		/// <param name="document"></param>
		/// <param name="id">A unique id that identifies an Image embedded in this document.</param>
		/// <param name="name">The name of this Picture.</param>
		/// <param name="descr">The description of this Picture.</param>
		static internal Picture CreatePicture(DocX document, string id, string name, string descr)
		{
			PackagePart part = document.package.GetPart(document.mainPart.GetRelationship(id).TargetUri);

			int newDocPrId = 1;
			List<string> existingIds = new List<string>();
			foreach (var bookmarkId in document.Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName)))
			{
				var idAtt = bookmarkId.Attributes().FirstOrDefault(x => x.Name.LocalName == "id");
				if (idAtt != null)
					existingIds.Add(idAtt.Value);
			}

			while (existingIds.Contains(newDocPrId.ToString()))
				newDocPrId++;


			int cx, cy;

			using (System.Drawing.Image img = System.Drawing.Image.FromStream(new PackagePartStream(part.GetStream())))
			{
				cx = img.Width * 9526;
				cy = img.Height * 9526;
			}

			XElement e = new XElement(DocX.w + "drawing");

			XElement xml = XElement.Parse
				 (string.Format(@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                            <wp:extent cx=""{0}"" cy=""{1}"" />
                            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                            <wp:docPr id=""{5}"" name=""{3}"" descr=""{4}"" />
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                    <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                        <pic:nvPicPr>
                                        <pic:cNvPr id=""0"" name=""{3}"" />
                                            <pic:cNvPicPr />
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed=""{2}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                                            <a:stretch>
                                                <a:fillRect />
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:off x=""0"" y=""0"" />
                                                <a:ext cx=""{0}"" cy=""{1}"" />
                                            </a:xfrm>
                                            <a:prstGeom prst=""rect"">
                                                <a:avLst />
                                            </a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing></w:r>
                    ", cx, cy, id, name, descr, newDocPrId.ToString()));

			return new Picture(document, xml, new Image(document, document.mainPart.GetRelationship(id)));
		}

		// Removed because it confusses the API.
		//public Picture InsertPicture(int index, string imageID)
		//{
		//    return InsertPicture(index, imageID, string.Empty, string.Empty);
		//}

		/// <summary>
		/// Creates an Edit either a ins or a del with the specified content and date
		/// </summary>
		/// <param name="t">The type of this edit (ins or del)</param>
		/// <param name="edit_time">The time stamp to use for this edit</param>
		/// <param name="content">The initial content of this edit</param>
		/// <returns></returns>
		internal static XElement CreateEdit(EditType t, DateTime edit_time, object content)
		{
			if (t == EditType.del)
			{
				foreach (object o in (IEnumerable<XElement>)content)
				{
					if (o is XElement)
					{
						XElement e = (o as XElement);
						IEnumerable<XElement> ts = e.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName));

						for (int i = 0; i < ts.Count(); i++)
						{
							XElement text = ts.ElementAt(i);
							text.ReplaceWith(new XElement(DocX.w + "delText", text.Attributes(), text.Value));
						}
					}
				}
			}

			return
			(
				new XElement(DocX.w + t.ToString(),
					new XAttribute(DocX.w + "id", 0),
					new XAttribute(DocX.w + "author", WindowsIdentity.GetCurrent().Name),
					new XAttribute(DocX.w + "date", edit_time),
				content)
			);
		}

		internal Run GetFirstRunEffectedByEdit(int index, EditType type = EditType.ins)
		{
			int len = HelperFunctions.GetText(Xml).Length;

			// Make sure we are looking within an acceptable index range.
			if (index < 0 || ((type == EditType.ins && index > len) || (type == EditType.del && index >= len)))
				throw new ArgumentOutOfRangeException();

			// Need some memory that can be updated by the recursive search for the XElement to Split.
			int count = 0;
			Run theOne = null;

			GetFirstRunEffectedByEditRecursive(Xml, index, ref count, ref theOne, type);

			return theOne;
		}

		internal void GetFirstRunEffectedByEditRecursive(XElement Xml, int index, ref int count, ref Run theOne, EditType type)
		{
			count += HelperFunctions.GetSize(Xml);

			// If the EditType is deletion then we must return the next blah
			if (count > 0 && ((type == EditType.del && count > index) || (type == EditType.ins && count >= index)))
			{
				// Correct the index
				foreach (XElement e in Xml.ElementsBeforeSelf())
					count -= HelperFunctions.GetSize(e);

				count -= HelperFunctions.GetSize(Xml);

				// We have found the element, now find the run it belongs to.
				while ((Xml.Name.LocalName != "r") && (Xml.Name.LocalName != "pPr"))
					Xml = Xml.Parent;

				theOne = new Run(Document, Xml, count);
				return;
			}

			if (Xml.HasElements)
				foreach (XElement e in Xml.Elements())
					if (theOne == null)
						GetFirstRunEffectedByEditRecursive(e, index, ref count, ref theOne, type);
		}

		/// <!-- 
		/// Bug found and fixed by krugs525 on August 12 2009.
		/// Use TFS compare to see exact code change.
		/// -->
		static internal int GetElementTextLength(XElement run)
		{
			int count = 0;

			if (run == null)
				return count;

			foreach (var d in run.Descendants())
			{
				switch (d.Name.LocalName)
				{
					case "tab":
						if (d.Parent.Name.LocalName != "tabs")
							goto case "br"; break;
					case "br": count++; break;
					case "t": goto case "delText";
					case "delText": count += d.Value.Length; break;
					default: break;
				}
			}
			return count;
		}

		internal XElement[] SplitEdit(XElement edit, int index, EditType type)
		{
			Run run = GetFirstRunEffectedByEdit(index, type);

			XElement[] splitRun = Run.SplitRun(run, index, type);

			XElement splitLeft = new XElement(edit.Name, edit.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
			if (GetElementTextLength(splitLeft) == 0)
				splitLeft = null;

			XElement splitRight = new XElement(edit.Name, edit.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
			if (GetElementTextLength(splitRight) == 0)
				splitRight = null;

			return
			(
				new XElement[]
				{
					splitLeft,
					splitRight
				}
			);
		}

		/// <summary>
		/// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
		/// </summary>
		/// <example>
		/// <code> 
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Create a text formatting.
		///     Formatting f = new Formatting();
		///     f.FontColor = Color.Red;
		///     f.Size = 30;
		///
		///     // Iterate through the Paragraphs in this document.
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
		///         p.InsertText("Start: ", true, f);
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <example>
		/// Inserting tabs using the \t switch.
		/// <code>  
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///      // Create a text formatting.
		///      Formatting f = new Formatting();
		///      f.FontColor = Color.Red;
		///      f.Size = 30;
		///        
		///      // Iterate through the paragraphs in this document.
		///      foreach (Paragraph p in document.Paragraphs)
		///      {
		///          // Insert the string "\tEnd" at the end of every paragraph and flag it as a change.
		///          p.InsertText("\tEnd", true, f);
		///      }
		///       
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="Paragraph.RemoveText(int, bool)"/>
		/// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		/// <param name="value">The System.String to insert.</param>
		/// <param name="trackChanges">Flag this insert as a change.</param>
		/// <param name="formatting">The text formatting.</param>
		public void InsertText(string value, bool trackChanges = false, Formatting formatting = null)
		{
			// Default values for optional parameters must be compile time constants.
			// Would have like to write 'public void InsertText(string value, bool trackChanges = false, Formatting formatting = new Formatting())
			if (formatting == null) formatting = new Formatting();

			List<XElement> newRuns = HelperFunctions.FormatInput(value, formatting.Xml);
			Xml.Add(newRuns);

			HelperFunctions.RenumberIDs(Document);
		}

		/// <summary>
		/// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
		/// </summary>
		/// <example>
		/// <code> 
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Create a text formatting.
		///     Formatting f = new Formatting();
		///     f.FontColor = Color.Red;
		///     f.Size = 30;
		///
		///     // Iterate through the Paragraphs in this document.
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
		///         p.InsertText(0, "Start: ", true, f);
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <example>
		/// Inserting tabs using the \t switch.
		/// <code>  
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Create a text formatting.
		///     Formatting f = new Formatting();
		///     f.FontColor = Color.Red;
		///     f.Size = 30;
		///
		///     // Iterate through the paragraphs in this document.
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
		///         p.InsertText(0, "\tStart:\t", true, f);
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="Paragraph.RemoveText(int, bool)"/>
		/// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		/// <param name="index">The index position of the insertion.</param>
		/// <param name="value">The System.String to insert.</param>
		/// <param name="trackChanges">Flag this insert as a change.</param>
		/// <param name="formatting">The text formatting.</param>
		public void InsertText(int index, string value, bool trackChanges = false, Formatting formatting = null)
		{
			// Timestamp to mark the start of insert
			DateTime now = DateTime.Now;
			DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

			// Get the first run effected by this Insert
			Run run = GetFirstRunEffectedByEdit(index);

			if (run == null)
			{
				object insert;
				if (formatting != null) //not sure how to get original formatting here when run == null
					insert = HelperFunctions.FormatInput(value, formatting.Xml);
				else
					insert = HelperFunctions.FormatInput(value, null);

				if (trackChanges)
					insert = CreateEdit(EditType.ins, insert_datetime, insert);
				Xml.Add(insert);
			}

			else
			{
				object newRuns;
				var rprel = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));
				if (formatting != null)
				{
					//merge 2 formattings properly
					Formatting finfmt = null;
					Formatting oldfmt = null;

					if (rprel != null)
					{
						oldfmt = Formatting.Parse(rprel);
					}

					if (oldfmt != null)
					{
						finfmt = oldfmt.Clone();
						if (formatting.Bold.HasValue) { finfmt.Bold = formatting.Bold; }
						if (formatting.CapsStyle.HasValue) { finfmt.CapsStyle = formatting.CapsStyle; }
						if (formatting.FontColor.HasValue) { finfmt.FontColor = formatting.FontColor; }
						finfmt.FontFamily = formatting.FontFamily;
						if (formatting.Hidden.HasValue) { finfmt.Hidden = formatting.Hidden; }
						if (formatting.Highlight.HasValue) { finfmt.Highlight = formatting.Highlight; }
						if (formatting.Italic.HasValue) { finfmt.Italic = formatting.Italic; }
						if (formatting.Kerning.HasValue) { finfmt.Kerning = formatting.Kerning; }
						finfmt.Language = formatting.Language;
						if (formatting.Misc.HasValue) { finfmt.Misc = formatting.Misc; }
						if (formatting.PercentageScale.HasValue) { finfmt.PercentageScale = formatting.PercentageScale; }
						if (formatting.Position.HasValue) { finfmt.Position = formatting.Position; }
						if (formatting.Script.HasValue) { finfmt.Script = formatting.Script; }
						if (formatting.Size.HasValue) { finfmt.Size = formatting.Size; }
						if (formatting.Spacing.HasValue) { finfmt.Spacing = formatting.Spacing; }
						if (formatting.StrikeThrough.HasValue) { finfmt.StrikeThrough = formatting.StrikeThrough; }
						if (formatting.UnderlineColor.HasValue) { finfmt.UnderlineColor = formatting.UnderlineColor; }
						if (formatting.UnderlineStyle.HasValue) { finfmt.UnderlineStyle = formatting.UnderlineStyle; }
					}
					else
					{
						finfmt = formatting;
					}

					newRuns = HelperFunctions.FormatInput(value, finfmt.Xml);
				}
				else
				{
					newRuns = HelperFunctions.FormatInput(value, rprel);
				}

				// The parent of this Run
				XElement parentElement = run.Xml.Parent;
				switch (parentElement.Name.LocalName)
				{
					case "ins":
						{
							// The datetime that this ins was created
							DateTime parent_ins_date = DateTime.Parse(parentElement.Attribute(XName.Get("date", DocX.w.NamespaceName)).Value);

							/* 
                             * Special case: You want to track changes,
                             * and the first Run effected by this insert
                             * has a datetime stamp equal to now.
                            */
							if (trackChanges && parent_ins_date.CompareTo(insert_datetime) == 0)
							{
								/*
                                 * Inserting into a non edit and this special case, is the same procedure.
                                */
								goto default;
							}

							/*
                             * If not the special case above, 
                             * then inserting into an ins or a del, is the same procedure.
                            */
							goto case "del";
						}

					case "del":
						{
							object insert = newRuns;
							if (trackChanges)
								insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

							// Split this Edit at the point you want to insert
							XElement[] splitEdit = SplitEdit(parentElement, index, EditType.ins);

							// Replace the origional run
							parentElement.ReplaceWith
							(
								splitEdit[0],
								insert,
								splitEdit[1]
							);

							break;
						}

					default:
						{
							object insert = newRuns;
							if (trackChanges && !parentElement.Name.LocalName.Equals("ins"))
								insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

							// Special case to deal with Page Number elements.
							//if (parentElement.Name.LocalName.Equals("fldSimple"))
							//    parentElement.AddBeforeSelf(insert);

							else
							{
								// Split this run at the point you want to insert
								XElement[] splitRun = Run.SplitRun(run, index);

								// Replace the origional run
								run.Xml.ReplaceWith
								(
									splitRun[0],
									insert,
									splitRun[1]
								);
							}

							break;
						}
				}
			}

			HelperFunctions.RenumberIDs(Document);
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <returns>This Paragraph in curent culture</returns>
		/// <example>
		/// Add a new Paragraph with russian text to this document and then set language of text to local culture.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph with russian text and set curent local culture to it.
		///     Paragraph p = document.InsertParagraph("Привет мир!").CurentCulture();
		///       
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph CurentCulture()
		{
			ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName),
				string.Empty,
				new XAttribute(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture.Name));
			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="culture">The CultureInfo for text</param>
		/// <returns>This Paragraph in curent culture</returns>
		/// <example>
		/// Add a new Paragraph with russian text to this document and then set language of text to local culture.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph with russian text and set specific culture to it.
		///     Paragraph p = document.InsertParagraph("Привет мир").Culture(CultureInfo.CreateSpecificCulture("ru-RU"));
		///       
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph Culture(CultureInfo culture)
		{
			ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName), string.Empty,
				new XAttribute(XName.Get("val", DocX.w.NamespaceName), culture.Name));
			return this;
		}

		/// <summary>
		/// Append text to this Paragraph.
		/// </summary>
		/// <param name="text">The text to append.</param>
		/// <returns>This Paragraph with the new text appened.</returns>
		/// <example>
		/// Add a new Paragraph to this document and then append some text to it.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph and Append some text to it.
		///     Paragraph p = document.InsertParagraph().Append("Hello World!!!");
		///       
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph Append(string text)
		{
			List<XElement> newRuns = HelperFunctions.FormatInput(text, null);
			Xml.Add(newRuns);

			this.runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(newRuns.Count()).ToList();

			return this;
		}

		/// <summary></summary>
		/// <param name="lineType"></param>
		/// <param name="size"></param>
		/// <param name="space"></param>
		/// <param name="color"></param>
		public void InsertHorizontalLine(string lineType = "single", int size = 6, int space = 1, string color = "auto")
		{
			var pPr = this.GetOrCreate_pPr();
			var border = pPr.Element(XName.Get("pBdr", DocX.w.NamespaceName));
			if (border == null)
			{
				pPr.Add(new XElement(XName.Get("pBdr", DocX.w.NamespaceName)));
				border = pPr.Element(XName.Get("pBdr", DocX.w.NamespaceName));
				border.Add(new XElement(XName.Get("bottom", DocX.w.NamespaceName)));
				var bottom = border.Element(XName.Get("bottom", DocX.w.NamespaceName));
				bottom.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), lineType);
				bottom.SetAttributeValue(XName.Get("sz", DocX.w.NamespaceName), size.ToString());
				bottom.SetAttributeValue(XName.Get("space", DocX.w.NamespaceName), space.ToString());
				bottom.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), color);
			}
		}

		/// <summary>
		/// Append a hyperlink to a Paragraph.
		/// </summary>
		/// <param name="h">The hyperlink to append.</param>
		/// <returns>The Paragraph with the hyperlink appended.</returns>
		/// <example>
		/// Creates a Paragraph with some text and a hyperlink.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///    // Add a hyperlink to this document.
		///    Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));
		///    
		///    // Add a new Paragraph to this document.
		///    Paragraph p = document.InsertParagraph();
		///    p.Append("My favourite search engine is ");
		///    p.AppendHyperlink(h);
		///    p.Append(", I think it's great.");
		///
		///    // Save all changes made to this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph AppendHyperlink(Hyperlink h)
		{
			// Convert the path of this mainPart to its equilivant rels file path.
			string path = mainPart.Uri.OriginalString.Replace("/word/", "");
			Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

			// Check to see if the rels file exists and create it if not.
			if (!Document.package.PartExists(rels_path))
				HelperFunctions.CreateRelsPackagePart(Document, rels_path);

			// Check to see if a rel for this Hyperlink exists, create it if not.
			var Id = GetOrGenerateRel(h);

			Xml.Add(h.Xml);
			Xml.Elements().Last().SetAttributeValue(DocX.r + "id", Id);

			this.runs = Xml.Elements().Last().Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();

			return this;
		}

		/// <summary>
		/// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		/// </summary>
		/// <param name="p">The Picture to append.</param>
		/// <returns>The Paragraph with the Picture now appended.</returns>
		/// <example>
		/// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		/// <code>
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///    // Add an image to the document. 
		///    Image     i = document.AddImage(@"Image.jpg");
		///    
		///    // Create a picture i.e. (A custom view of an image)
		///    Picture   p = i.CreatePicture();
		///    p.FlipHorizontal = true;
		///    p.Rotation = 10;
		///
		///    // Create a new Paragraph.
		///    Paragraph par = document.InsertParagraph();
		///    
		///    // Append content to the Paragraph.
		///    par.Append("Here is a cool picture")
		///       .AppendPicture(p)
		///       .Append(" don't you think so?");
		///
		///    // Save all changes made to this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph AppendPicture(Picture p)
		{
			// Convert the path of this mainPart to its equilivant rels file path.
			string path = mainPart.Uri.OriginalString.Replace("/word/", "");
			Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

			// Check to see if the rels file exists and create it if not.
			if (!Document.package.PartExists(rels_path))
				HelperFunctions.CreateRelsPackagePart(Document, rels_path);

			// Check to see if a rel for this Picture exists, create it if not.
			var Id = GetOrGenerateRel(p);

			// Add the Picture Xml to the end of the Paragragraph Xml.
			Xml.Add(p.Xml);

			// Extract the attribute id from the Pictures Xml.
			XAttribute a_id =
			(
				from e in Xml.Elements().Last().Descendants()
				where e.Name.LocalName.Equals("blip")
				select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
			).Single();

			// Set its value to the Pictures relationships id.
			a_id.SetValue(Id);

			// For formatting such as .Bold()
			this.runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(p.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Count()).ToList();

			return this;
		}

		/// <summary>
		/// Add an equation to a document.
		/// </summary>
		/// <param name="equation">The Equation to append.</param>
		/// <returns>The Paragraph with the Equation now appended.</returns>
		/// <example>
		/// Add an equation to a document.
		/// <code>
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///    // Add an equation to the document. 
		///    document.AddEquation("x=y+z");
		///    
		///    // Save all changes made to this document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph AppendEquation(String equation)
		{
			// Create equation element
			XElement oMathPara =
				new XElement
				(
					XName.Get("oMathPara", DocX.m.NamespaceName),
					new XElement
					(
						XName.Get("oMath", DocX.m.NamespaceName),
						new XElement
						(
							XName.Get("r", DocX.w.NamespaceName),
							new Formatting() { FontFamily = new Font("Cambria Math") }.Xml,    // create formatting
							new XElement(XName.Get("t", DocX.m.NamespaceName), equation)                            // create equation string
						)
					)
				);

			// Add equation element into paragraph xml and update runs collection
			Xml.Add(oMathPara);
			runs = Xml.Elements(XName.Get("oMathPara", DocX.m.NamespaceName)).ToList();

			// Return paragraph with equation
			return this;
		}

		/// <summary></summary>
		/// <param name="bookmarkName"></param>
		/// <returns></returns>
		public bool ValidateBookmark(string bookmarkName)
		{
			return GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
		}

		/// <summary></summary>
		/// <param name="bookmarkName"></param>
		/// <returns></returns>
		public Paragraph AppendBookmark(String bookmarkName)
		{
			XElement wBookmarkStart = new XElement(
				XName.Get("bookmarkStart", DocX.w.NamespaceName),
				new XAttribute(XName.Get("id", DocX.w.NamespaceName), 0),
				new XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName));
			Xml.Add(wBookmarkStart);

			XElement wBookmarkEnd = new XElement(
				XName.Get("bookmarkEnd", DocX.w.NamespaceName),
				new XAttribute(XName.Get("id", DocX.w.NamespaceName), 0),
				new XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName));
			Xml.Add(wBookmarkEnd);

			return this;
		}

		/// <summary></summary>
		/// <returns></returns>
		public IEnumerable<Bookmark> GetBookmarks()
		{
			return Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
						.Select(x => x.Attribute(XName.Get("name", DocX.w.NamespaceName)))
						.Select(x => new Bookmark
						{
							Name = x.Value,
							Paragraph = this
						});
		}

		/// <summary></summary>
		/// <param name="toInsert"></param>
		/// <param name="bookmarkName"></param>
		public void InsertAtBookmark(string toInsert, string bookmarkName)
		{
			var bookmark = Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
								.Where(x => x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value == bookmarkName).SingleOrDefault();
			if (bookmark != null)
			{

				var run = HelperFunctions.FormatInput(toInsert, null);
				bookmark.AddBeforeSelf(run);
				runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
				HelperFunctions.RenumberIDs(Document);
			}
		}

		/// <summary></summary>
		/// <param name="toInsert"></param>
		/// <param name="bookmarkName"></param>
		public void ReplaceAtBookmark(string toInsert, string bookmarkName)
		{
			XElement bookmark = Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
				.Where(x => x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value == bookmarkName)
				.SingleOrDefault();
			if (bookmark == null)
				return;

			XNode nextNode = bookmark.NextNode;
			XElement nextElement = nextNode as XElement;
			while (nextElement == null || nextElement.Name.NamespaceName != DocX.w.NamespaceName || (nextElement.Name.LocalName != "r" && nextElement.Name.LocalName != "bookmarkEnd"))
			{
				nextNode = nextNode.NextNode;
				nextElement = nextNode as XElement;
			}

			// Check if next element is a bookmarkEnd
			if (nextElement.Name.LocalName == "bookmarkEnd")
			{
				ReplaceAtBookmark_Add(toInsert, bookmark);
				return;
			}

			XElement contentElement = nextElement.Elements(XName.Get("t", DocX.w.NamespaceName)).FirstOrDefault();
			if (contentElement == null)
			{
				ReplaceAtBookmark_Add(toInsert, bookmark);
				return;
			}

			contentElement.Value = toInsert;
		}

		private void ReplaceAtBookmark_Add(string toInsert, XElement bookmark)
		{
			var run = HelperFunctions.FormatInput(toInsert, null);
			bookmark.AddAfterSelf(run);
			runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
			HelperFunctions.RenumberIDs(Document);
		}


		internal string GetOrGenerateRel(Picture p)
		{
			string image_uri_string = p.img.pr.TargetUri.OriginalString;

			// Search for a relationship with a TargetUri that points at this Image.
			var Id =
			(
				from r in mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				where r.TargetUri.OriginalString == image_uri_string
				select r.Id
			).SingleOrDefault();

			// If such a relation dosen't exist, create one.
			if (Id == null)
			{
				// Check to see if a relationship for this Picture exists and create it if not.
				PackageRelationship pr = mainPart.CreateRelationship(p.img.pr.TargetUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				Id = pr.Id;
			}
			return Id;
		}

		internal string GetOrGenerateRel(Hyperlink h)
		{
			string image_uri_string = h.Uri.OriginalString;

			// Search for a relationship with a TargetUri that points at this Image.
			var Id =
			(
				from r in mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
				where r.TargetUri.OriginalString == image_uri_string
				select r.Id
			).SingleOrDefault();

			// If such a relation dosen't exist, create one.
			if (Id == null)
			{
				// Check to see if a relationship for this Picture exists and create it if not.
				PackageRelationship pr = mainPart.CreateRelationship(h.Uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
				Id = pr.Id;
			}
			return Id;
		}

		/// <summary>
		/// Insert a Picture into a Paragraph at the given text index.
		/// If not index is provided defaults to 0.
		/// </summary>
		/// <param name="p">The Picture to insert.</param>
		/// <param name="index">The text index to insert at.</param>
		/// <returns>The modified Paragraph.</returns>
		/// <example>
		/// <code>
		///Load test document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Add Headers and Footers into this document.
		///    document.AddHeaders();
		///    document.AddFooters();
		///    document.DifferentFirstPage = true;
		///    document.DifferentOddAndEvenPages = true;
		///
		///    // Add an Image to this document.
		///    Novacode.Image img = document.AddImage(directory_documents + "purple.png");
		///
		///    // Create a Picture from this Image.
		///    Picture pic = img.CreatePicture();
		///
		///    // Main document.
		///    Paragraph p0 = document.InsertParagraph("Hello");
		///    p0.InsertPicture(pic, 3);
		///
		///    // Header first.
		///    Paragraph p1 = document.Headers.first.InsertParagraph("----");
		///    p1.InsertPicture(pic, 2);
		///
		///    // Header odd.
		///    Paragraph p2 = document.Headers.odd.InsertParagraph("----");
		///    p2.InsertPicture(pic, 2);
		///
		///    // Header even.
		///    Paragraph p3 = document.Headers.even.InsertParagraph("----");
		///    p3.InsertPicture(pic, 2);
		///
		///    // Footer first.
		///    Paragraph p4 = document.Footers.first.InsertParagraph("----");
		///    p4.InsertPicture(pic, 2);
		///
		///    // Footer odd.
		///    Paragraph p5 = document.Footers.odd.InsertParagraph("----");
		///    p5.InsertPicture(pic, 2);
		///
		///    // Footer even.
		///    Paragraph p6 = document.Footers.even.InsertParagraph("----");
		///    p6.InsertPicture(pic, 2);
		///
		///    // Save this document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public Paragraph InsertPicture(Picture p, int index = 0)
		{
			// Convert the path of this mainPart to its equilivant rels file path.
			string path = mainPart.Uri.OriginalString.Replace("/word/", "");
			Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

			// Check to see if the rels file exists and create it if not.
			if (!Document.package.PartExists(rels_path))
				HelperFunctions.CreateRelsPackagePart(Document, rels_path);

			// Check to see if a rel for this Picture exists, create it if not.
			var Id = GetOrGenerateRel(p);

			XElement p_xml;
			if (index == 0)
			{
				// Add this hyperlink as the last element.
				Xml.AddFirst(p.Xml);

				// Extract the picture back out of the DOM.
				p_xml = (XElement)Xml.FirstNode;
			}

			else
			{
				// Get the first run effected by this Insert
				Run run = GetFirstRunEffectedByEdit(index);

				if (run == null)
				{
					// Add this picture as the last element.
					Xml.Add(p.Xml);

					// Extract the picture back out of the DOM.
					p_xml = (XElement)Xml.LastNode;
				}

				else
				{
					// Split this run at the point you want to insert
					XElement[] splitRun = Run.SplitRun(run, index);

					// Replace the origional run.
					run.Xml.ReplaceWith
					(
						splitRun[0],
						p.Xml,
						splitRun[1]
					);

					// Get the first run effected by this Insert
					run = GetFirstRunEffectedByEdit(index);

					// The picture has to be the next element, extract it back out of the DOM.
					p_xml = (XElement)run.Xml.NextNode;
				}
			}
			// Extract the attribute id from the Pictures Xml.
			XAttribute a_id =
			(
				from e in p_xml.Descendants()
				where e.Name.LocalName.Equals("blip")
				select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
			).Single();

			// Set its value to the Pictures relationships id.
			a_id.SetValue(Id);


			return this;
		}

		/// <summary>
		/// Append text on a new line to this Paragraph.
		/// </summary>
		/// <param name="text">The text to append.</param>
		/// <returns>This Paragraph with the new text appened.</returns>
		/// <example>
		/// Add a new Paragraph to this document and then append a new line with some text to it.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph and Append a new line with some text to it.
		///     Paragraph p = document.InsertParagraph().AppendLine("Hello World!!!");
		///       
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph AppendLine(string text)
		{
			return Append("\n" + text);
		}

		/// <summary>
		/// Append a new line to this Paragraph.
		/// </summary>
		/// <returns>This Paragraph with a new line appeneded.</returns>
		/// <example>
		/// Add a new Paragraph to this document and then append a new line to it.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph and Append a new line with some text to it.
		///     Paragraph p = document.InsertParagraph().AppendLine();
		///       
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public Paragraph AppendLine()
		{
			return Append("\n");
		}

		internal void ApplyTextFormattingProperty(XName textFormatPropName, string value, object content)
		{
			XElement rPr = null;

			if (runs.Count == 0)
			{
				XElement pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
				if (pPr == null)
				{
					Xml.AddFirst(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
					pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
				}

				rPr = pPr.Element(XName.Get("rPr", DocX.w.NamespaceName));
				if (rPr == null)
				{
					pPr.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
					rPr = pPr.Element(XName.Get("rPr", DocX.w.NamespaceName));
				}

				rPr.SetElementValue(textFormatPropName, value);
				var last = rPr.Elements(textFormatPropName).Last();
				if (content as XAttribute != null)//If content is an attribute
				{
					if (last.Attribute(((XAttribute)(content)).Name) == null)
					{
						last.Add(content); //Add this attribute if element doesn't have it
					}
					else
					{
						last.Attribute(((XAttribute)(content)).Name).Value = ((XAttribute)(content)).Value; //Apply value only if element already has it
					}
				}
				return;
			}

			var contentIsListOfFontProperties = false;
			var fontProps = content as IEnumerable;
			if (fontProps != null)
			{
				foreach (object property in fontProps)
				{
					contentIsListOfFontProperties = (property as XAttribute != null);
				}
			}

			foreach (XElement run in runs)
			{
				rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
				if (rPr == null)
				{
					run.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
					rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
				}

				rPr.SetElementValue(textFormatPropName, value);
				XElement last = rPr.Elements(textFormatPropName).Last();

				if (contentIsListOfFontProperties) //if content is a list of attributes, as in the case when specifying a font family 
				{
					foreach (object property in fontProps)
					{
						if (last.Attribute(((XAttribute)(property)).Name) == null)
						{
							last.Add(property); //Add this attribute if element doesn't have it
						}
						else
						{
							last.Attribute(((XAttribute)(property)).Name).Value =
							  ((XAttribute)(property)).Value; //Apply value only if element already has it
						}
					}
				}

				if (content as XAttribute != null)//If content is an attribute
				{
					if (last.Attribute(((XAttribute)(content)).Name) == null)
					{
						last.Add(content); //Add this attribute if element doesn't have it
					}
					else
					{
						last.Attribute(((XAttribute)(content)).Name).Value = ((XAttribute)(content)).Value; //Apply value only if element already has it
					}
				}
				else
				{
					//IMPORTANT
					//But what to do if it is not?
				}
			}
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <returns>This Paragraph with the last appended text bold.</returns>
		/// <example>
		/// Append text to this Paragraph and then make it bold.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Bold").Bold()
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Bold()
		{
			ApplyTextFormattingProperty(XName.Get("b", DocX.w.NamespaceName), string.Empty, null);
			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <returns>This Paragraph with the last appended text italic.</returns>
		/// <example>
		/// Append text to this Paragraph and then make it italic.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Italic").Italic()
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Italic()
		{
			ApplyTextFormattingProperty(XName.Get("i", DocX.w.NamespaceName), string.Empty, null);
			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="c">A color to use on the appended text.</param>
		/// <returns>This Paragraph with the last appended text colored.</returns>
		/// <example>
		/// Append text to this Paragraph and then color it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Blue").Color(Color.Blue)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Color(Color c)
		{
			ApplyTextFormattingProperty(XName.Get("color", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), c.ToHex()));
			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="underlineStyle">The underline style to use for the appended text.</param>
		/// <returns>This Paragraph with the last appended text underlined.</returns>
		/// <example>
		/// Append text to this Paragraph and then underline it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Underlined").UnderlineStyle(UnderlineStyle.doubleLine)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph UnderlineStyle(UnderlineStyle underlineStyle)
		{
			string value;
			switch (underlineStyle)
			{
				case Novacode.UnderlineStyle.none: value = string.Empty; break;
				case Novacode.UnderlineStyle.singleLine: value = "single"; break;
				case Novacode.UnderlineStyle.doubleLine: value = "double"; break;
				default: value = underlineStyle.ToString(); break;
			}

			ApplyTextFormattingProperty(XName.Get("u", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
			return this;
		}

		private Table followingTable;

		///<summary>
		/// Returns table following the paragraph. Null if the following element isn't table.
		///</summary>
		public Table FollowingTable
		{
			get
			{
				return followingTable;
			}
			internal set
			{
				followingTable = value;
			}
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="fontSize">The font size to use for the appended text.</param>
		/// <returns>This Paragraph with the last appended text resized.</returns>
		/// <example>
		/// Append text to this Paragraph and then resize it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Big").FontSize(20)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph FontSize(double fontSize)
		{
			double temp = fontSize * 2;

			if (temp - (int)temp == 0)
			{
				if (!(fontSize > 0 && fontSize < 1639))
					throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
			}

			else
				throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");

			ApplyTextFormattingProperty(XName.Get("sz", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize * 2));
			ApplyTextFormattingProperty(XName.Get("szCs", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize * 2));

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="fontName">The font to use for the appended text.</param>
		/// <returns>This Paragraph with the last appended text's font changed.</returns>
		public Paragraph Font(string fontName)
		{
			return Font(new Font(fontName));
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="fontFamily">The font to use for the appended text.</param>
		/// <returns>This Paragraph with the last appended text's font changed.</returns>
		/// <example>
		/// Append text to this Paragraph and then change its font.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Times new roman").Font(new FontFamily("Times new roman"))
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Font(Font fontFamily)
		{
			ApplyTextFormattingProperty
			(
				XName.Get("rFonts", DocX.w.NamespaceName),
				string.Empty,
				new[]
				{
					new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name),
					new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name), // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                    new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name),    // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                    new XAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), fontFamily.Name) // DOCX in china #57
                }
			);

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="capsStyle">The caps style to apply to the last appended text.</param>
		/// <returns>This Paragraph with the last appended text's caps style changed.</returns>
		/// <example>
		/// Append text to this Paragraph and then set it to full caps.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("Capitalized").CapsStyle(CapsStyle.caps)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph CapsStyle(CapsStyle capsStyle)
		{
			switch (capsStyle)
			{
				case Novacode.CapsStyle.none:
					break;

				default:
					{
						ApplyTextFormattingProperty(XName.Get(capsStyle.ToString(), DocX.w.NamespaceName), string.Empty, null);
						break;
					}
			}

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="script">The script style to apply to the last appended text.</param>
		/// <returns>This Paragraph with the last appended text's script style changed.</returns>
		/// <example>
		/// Append text to this Paragraph and then set it to superscript.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("superscript").Script(Script.superscript)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Script(Script script)
		{
			switch (script)
			{
				case Novacode.Script.none:
					break;

				default:
					{
						ApplyTextFormattingProperty(XName.Get("vertAlign", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), script.ToString()));
						break;
					}
			}

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		///<param name="highlight">The highlight to apply to the last appended text.</param>
		/// <returns>This Paragraph with the last appended text highlighted.</returns>
		/// <example>
		/// Append text to this Paragraph and then highlight it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("highlighted").Highlight(Highlight.green)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Highlight(Highlight highlight)
		{
			switch (highlight)
			{
				case Novacode.Highlight.none:
					break;

				default:
					{
						ApplyTextFormattingProperty(XName.Get("highlight", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight.ToString()));
						break;
					}
			}

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="misc">The miscellaneous property to set.</param>
		/// <returns>This Paragraph with the last appended text changed by a miscellaneous property.</returns>
		/// <example>
		/// Append text to this Paragraph and then apply a miscellaneous property.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("outlined").Misc(Misc.outline)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Misc(Misc misc)
		{
			switch (misc)
			{
				case Novacode.Misc.none:
					break;

				case Novacode.Misc.outlineShadow:
					{
						ApplyTextFormattingProperty(XName.Get("outline", DocX.w.NamespaceName), string.Empty, null);
						ApplyTextFormattingProperty(XName.Get("shadow", DocX.w.NamespaceName), string.Empty, null);

						break;
					}

				case Novacode.Misc.engrave:
					{
						ApplyTextFormattingProperty(XName.Get("imprint", DocX.w.NamespaceName), string.Empty, null);

						break;
					}

				default:
					{
						ApplyTextFormattingProperty(XName.Get(misc.ToString(), DocX.w.NamespaceName), string.Empty, null);

						break;
					}
			}

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="strikeThrough">The strike through style to used on the last appended text.</param>
		/// <returns>This Paragraph with the last appended text striked.</returns>
		/// <example>
		/// Append text to this Paragraph and then strike it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("striked").StrikeThrough(StrikeThrough.doubleStrike)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph StrikeThrough(StrikeThrough strikeThrough)
		{
			string value;
			switch (strikeThrough)
			{
				case Novacode.StrikeThrough.strike: value = "strike"; break;
				case Novacode.StrikeThrough.doubleStrike: value = "dstrike"; break;
				default: return this;
			}

			ApplyTextFormattingProperty(XName.Get(value, DocX.w.NamespaceName), string.Empty, null);

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <param name="underlineColor">The underline color to use, if no underline is set, a single line will be used.</param>
		/// <returns>This Paragraph with the last appended text underlined in a color.</returns>
		/// <example>
		/// Append text to this Paragraph and then underline it using a color.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("color underlined").UnderlineStyle(UnderlineStyle.dotted).UnderlineColor(Color.Orange)
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph UnderlineColor(Color underlineColor)
		{
			foreach (XElement run in runs)
			{
				XElement rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
				if (rPr == null)
				{
					run.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
					rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
				}

				XElement u = rPr.Element(XName.Get("u", DocX.w.NamespaceName));
				if (u == null)
				{
					rPr.SetElementValue(XName.Get("u", DocX.w.NamespaceName), string.Empty);
					u = rPr.Element(XName.Get("u", DocX.w.NamespaceName));
					u.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "single");
				}

				u.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), underlineColor.ToHex());
			}

			return this;
		}

		/// <summary>
		/// For use with Append() and AppendLine()
		/// </summary>
		/// <returns>This Paragraph with the last appended text hidden.</returns>
		/// <example>
		/// Append text to this Paragraph and then hide it.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Paragraph.
		///     Paragraph p = document.InsertParagraph();
		///
		///     p.Append("I am ")
		///     .Append("hidden").Hide()
		///     .Append(" I am not");
		///        
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public Paragraph Hide()
		{
			ApplyTextFormattingProperty(XName.Get("vanish", DocX.w.NamespaceName), string.Empty, null);

			return this;
		}

		/// <summary></summary>
		public float LineSpacing
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));

				if (spacing != null)
				{
					XAttribute line = spacing.Attribute(XName.Get("line", DocX.w.NamespaceName));
					if (line != null)
					{
						float f;

						if (float.TryParse(line.Value, out f))
							return f / 20.0f;
					}
				}

				return 1.1f * 20.0f;
			}

			set
			{
				Spacing(value);
			}
		}


		/// <summary>
		/// Set the linespacing for this paragraph manually.
		/// </summary>
		/// <param name="spacingType">The type of spacing to be set, can be either Before, After or Line (Standard line spacing).</param>
		/// <param name="spacingFloat">A float value of the amount of spacing. Equals the value that van be set in Word using the "Line and Paragraph spacing" button.</param>
		public void SetLineSpacing(LineSpacingType spacingType, float spacingFloat)
		{
			spacingFloat = spacingFloat * 240;
			int spacingValue = (int)spacingFloat;

			var pPr = this.GetOrCreate_pPr();
			var spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
			if (spacing == null)
			{
				pPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName)));
				spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
			}

			string spacingTypeAttribute = "";
			switch (spacingType)
			{
				case LineSpacingType.Line:
					{
						spacingTypeAttribute = "line";
						break;
					}
				case LineSpacingType.Before:
					{
						spacingTypeAttribute = "before";
						break;
					}
				case LineSpacingType.After:
					{
						spacingTypeAttribute = "after";
						break;
					}

			}

			spacing.SetAttributeValue(XName.Get(spacingTypeAttribute, DocX.w.NamespaceName), spacingValue);
		}

		/// <summary>
		/// Set the linespacing for this paragraph using the Auto value.
		/// </summary>
		/// <param name="spacingType">The type of spacing to be set automatically. Using Auto will set both Before and After. None will remove any linespacing.</param>
		public void SetLineSpacing(LineSpacingTypeAuto spacingType)
		{
			int spacingValue = 100;

			var pPr = this.GetOrCreate_pPr();
			var spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));

			if (spacingType.Equals(LineSpacingTypeAuto.None))
			{
				if (spacing != null)
				{
					spacing.Remove();
				}
			}

			else
			{

				if (spacing == null)
				{
					pPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName)));
					spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
				}

				string spacingTypeAttribute = "";
				string autoSpacingTypeAttribute = "";
				switch (spacingType)
				{
					case LineSpacingTypeAuto.AutoBefore:
						{
							spacingTypeAttribute = "before";
							autoSpacingTypeAttribute = "beforeAutospacing";
							break;
						}
					case LineSpacingTypeAuto.AutoAfter:
						{
							spacingTypeAttribute = "after";
							autoSpacingTypeAttribute = "afterAutospacing";
							break;
						}
					case LineSpacingTypeAuto.Auto:
						{
							spacingTypeAttribute = "before";
							autoSpacingTypeAttribute = "beforeAutospacing";
							spacing.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), spacingValue);
							spacing.SetAttributeValue(XName.Get("afterAutospacing", DocX.w.NamespaceName), 1);
							break;
						}

				}

				spacing.SetAttributeValue(XName.Get(autoSpacingTypeAttribute, DocX.w.NamespaceName), 1);
				spacing.SetAttributeValue(XName.Get(spacingTypeAttribute, DocX.w.NamespaceName), spacingValue);

			}

		}

		/// <summary></summary>
		/// <param name="spacing"></param>
		/// <returns></returns>
		public Paragraph Spacing(double spacing)
		{
			spacing *= 20;

			if (spacing - (int)spacing == 0)
			{
				if (!(spacing > -1585 && spacing < 1585))
					throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
			}

			else
				throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");

			ApplyTextFormattingProperty(XName.Get("spacing", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing));

			return this;
		}

		/// <summary></summary>
		/// <param name="spacingBefore"></param>
		/// <returns></returns>
		public Paragraph SpacingBefore(double spacingBefore)
		{
			spacingBefore *= 20;

			var pPr = GetOrCreate_pPr();
			var spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
			if (spacingBefore > 0)
			{
				if (spacing == null)
				{
					spacing = new XElement(XName.Get("spacing", DocX.w.NamespaceName));
					pPr.Add(spacing);
				}
				var attr = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName));
				if (attr == null)
					spacing.SetAttributeValue(XName.Get("before", DocX.w.NamespaceName), spacingBefore);
				else
					attr.SetValue(spacingBefore);
			}
			if (Math.Abs(spacingBefore) < 0.1f && spacing != null)
			{
				var attr = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName));
				attr.Remove();
				if (!spacing.HasAttributes)
					spacing.Remove();
			}

			return this;
		}

		/// <summary></summary>
		/// <param name="spacingAfter"></param>
		/// <returns></returns>
		public Paragraph SpacingAfter(double spacingAfter)
		{
			spacingAfter *= 20;

			var pPr = GetOrCreate_pPr();
			var spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
			if (spacingAfter > 0)
			{
				if (spacing == null)
				{
					spacing = new XElement(XName.Get("spacing", DocX.w.NamespaceName));
					pPr.Add(spacing);
				}
				var attr = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName));
				if (attr == null)
					spacing.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), spacingAfter);
				else
					attr.SetValue(spacingAfter);
			}
			if (Math.Abs(spacingAfter) < 0.1f && spacing != null)
			{
				var attr = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName));
				attr.Remove();
				if (!spacing.HasAttributes)
					spacing.Remove();
			}
			//ApplyTextFormattingProperty(XName.Get("after", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacingAfter));

			return this;
		}

		/// <summary></summary>
		/// <param name="kerning"></param>
		/// <returns></returns>
		public Paragraph Kerning(int kerning)
		{
			if (!new int?[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains(kerning)) throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
			ApplyTextFormattingProperty(XName.Get("kern", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning * 2));
			return this;
		}

		/// <summary></summary>
		/// <param name="position"></param>
		/// <returns></returns>
		public Paragraph Position(double position)
		{
			if (!(position > -1585 && position < 1585)) throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
			ApplyTextFormattingProperty(XName.Get("position", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), position * 2));
			return this;
		}

		/// <summary></summary>
		/// <param name="percentageScale"></param>
		/// <returns></returns>
		public Paragraph PercentageScale(int percentageScale)
		{
			if (!(new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(percentageScale)) throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
			ApplyTextFormattingProperty(XName.Get("w", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale));
			return this;
		}

		/// <summary>
		/// Append a field of type document property, this field will display the custom property cp, at the end of this paragraph.
		/// </summary>
		/// <param name="cp">The custom property to display.</param>
		/// <param name="trackChanges"></param>
		/// <param name="f">The formatting to use for this text.</param>
		/// <example>
		/// Create, add and display a custom property in a document.
		/// <code>
		/// // Load a document.
		///using (DocX document = DocX.Create("CustomProperty_Add.docx"))
		///{
		///    // Add a few Custom Properties to this document.
		///    document.AddCustomProperty(new CustomProperty("fname", "cathal"));
		///    document.AddCustomProperty(new CustomProperty("age", 24));
		///    document.AddCustomProperty(new CustomProperty("male", true));
		///    document.AddCustomProperty(new CustomProperty("newyear2012", new DateTime(2012, 1, 1)));
		///    document.AddCustomProperty(new CustomProperty("fav_num", 3.141592));
		///
		///    // Insert a new Paragraph and append a load of DocProperties.
		///    Paragraph p = document.InsertParagraph("fname: ")
		///        .AppendDocProperty(document.CustomProperties["fname"])
		///        .Append(", age: ")
		///        .AppendDocProperty(document.CustomProperties["age"])
		///        .Append(", male: ")
		///        .AppendDocProperty(document.CustomProperties["male"])
		///        .Append(", newyear2012: ")
		///        .AppendDocProperty(document.CustomProperties["newyear2012"])
		///        .Append(", fav_num: ")
		///        .AppendDocProperty(document.CustomProperties["fav_num"]);
		///    
		///    // Save the changes to the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public Paragraph AppendDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
		{
			this.InsertDocProperty(cp, trackChanges, f);
			return this;
		}

		/// <summary>
		/// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
		/// </summary>
		/// <param name="cp">The custom property to display.</param>
		/// <param name="trackChanges"></param>
		/// <param name="f">The formatting to use for this text.</param>
		/// <example>
		/// Create, add and display a custom property in a document.
		/// <code>
		/// // Load a document
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Create a custom property.
		///     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
		///        
		///     // Add this custom property to this document.
		///     document.AddCustomProperty(name);
		///
		///     // Create a text formatting.
		///     Formatting f = new Formatting();
		///     f.Bold = true;
		///     f.Size = 14;
		///     f.StrikeThrough = StrickThrough.strike;
		///
		///     // Insert a new paragraph.
		///     Paragraph p = document.InsertParagraph("Author: ", false, f);
		///
		///     // Insert a field of type document property to display the custom property name and track this change.
		///     p.InsertDocProperty(name, true, f);
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public DocProperty InsertDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
		{
			XElement f_xml = null;
			if (f != null)
				f_xml = f.Xml;

			XElement e = new XElement
			(
				XName.Get("fldSimple", DocX.w.NamespaceName),
				new XAttribute(XName.Get("instr", DocX.w.NamespaceName), string.Format(@"DOCPROPERTY {0} \* MERGEFORMAT", cp.Name)),
					new XElement(XName.Get("r", DocX.w.NamespaceName),
						new XElement(XName.Get("t", DocX.w.NamespaceName), f_xml, cp.Value))
			);

			XElement xml = e;
			if (trackChanges)
			{
				DateTime now = DateTime.Now;
				DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
				e = CreateEdit(EditType.ins, insert_datetime, e);
			}

			Xml.Add(e);

			return new DocProperty(Document, xml);
		}

		/// <summary>
		/// Removes characters from a Novacode.DocX.Paragraph.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Iterate through the paragraphs
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         // Remove the first two characters from every paragraph
		///         p.RemoveText(0, 2, false);
		///     }
		///        
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		/// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		/// <param name="index">The position to begin deleting characters.</param>
		/// <param name="count">The number of characters to delete</param>
		/// <param name="trackChanges">Track changes</param>
		public void RemoveText(int index, int count, bool trackChanges = false)
		{
			// Timestamp to mark the start of insert
			DateTime now = DateTime.Now;
			DateTime remove_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

			// The number of characters processed so far
			int processed = 0;

			do
			{
				// Get the first run effected by this Remove
				Run run = GetFirstRunEffectedByEdit(index, EditType.del);

				// The parent of this Run
				XElement parentElement = run.Xml.Parent;
				switch (parentElement.Name.LocalName)
				{
					case "ins":
						{
							XElement[] splitEditBefore = SplitEdit(parentElement, index, EditType.del);
							int min = Math.Min(count - processed, run.Xml.ElementsAfterSelf().Sum(e => GetElementTextLength(e)));
							XElement[] splitEditAfter = SplitEdit(parentElement, index + min, EditType.del);

							XElement temp = SplitEdit(splitEditBefore[1], index + min, EditType.del)[0];
							object middle = CreateEdit(EditType.del, remove_datetime, temp.Elements());
							processed += GetElementTextLength(middle as XElement);

							if (!trackChanges)
								middle = null;

							parentElement.ReplaceWith
							(
								splitEditBefore[0],
								middle,
								splitEditAfter[1]
							);

							processed += GetElementTextLength(middle as XElement);
							break;
						}

					case "del":
						{
							if (trackChanges)
							{
								// You cannot delete from a deletion, advance processed to the end of this del
								processed += GetElementTextLength(parentElement);
							}

							else
								goto case "ins";

							break;
						}

					default:
						{
							XElement[] splitRunBefore = Run.SplitRun(run, index, EditType.del);
							//int min = Math.Min(index + processed + (count - processed), run.EndIndex);
							int min = Math.Min(index + (count - processed), run.EndIndex);
							XElement[] splitRunAfter = Run.SplitRun(run, min, EditType.del);

							object middle = CreateEdit(EditType.del, remove_datetime, new List<XElement>() { Run.SplitRun(new Run(Document, splitRunBefore[1], run.StartIndex + GetElementTextLength(splitRunBefore[0])), min, EditType.del)[0] });
							processed += GetElementTextLength(middle as XElement);

							if (!trackChanges)
								middle = null;

							run.Xml.ReplaceWith
							(
								splitRunBefore[0],
								middle,
								splitRunAfter[1]
							);

							break;
						}
				}

				// If after this remove the parent element is empty, remove it.
				if (GetElementTextLength(parentElement) == 0)
				{
					if (parentElement.Parent != null && parentElement.Parent.Name.LocalName != "tc")
					{
						// Need to make sure there is no drawing element within the parent element.
						// Picture elements contain no text length but they are still content.
						if (parentElement.Descendants(XName.Get("drawing", DocX.w.NamespaceName)).Count() == 0)
							parentElement.Remove();
					}
				}
			}
			while (processed < count);

			HelperFunctions.RenumberIDs(Document);
		}


		/// <summary>
		/// Removes characters from a Novacode.DocX.Paragraph.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Iterate through the paragraphs
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         // Remove all but the first 2 characters from this Paragraph.
		///         p.RemoveText(2, false);
		///     }
		///        
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		/// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		/// <param name="index">The position to begin deleting characters.</param>
		/// <param name="trackChanges">Track changes</param>
		public void RemoveText(int index, bool trackChanges = false)
		{
			RemoveText(index, Text.Length - index, trackChanges);
		}

		/// <summary>
		/// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
		/// </summary>
		/// <example>
		/// <code>
		/// // Load a document using a relative filename.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // The formatting to match.
		///     Formatting matchFormatting = new Formatting();
		///     matchFormatting.Size = 10;
		///     matchFormatting.Italic = true;
		///     matchFormatting.FontFamily = new FontFamily("Times New Roman");
		///
		///     // The formatting to apply to the inserted text.
		///     Formatting newFormatting = new Formatting();
		///     newFormatting.Size = 22;
		///     newFormatting.UnderlineStyle = UnderlineStyle.dotted;
		///     newFormatting.Bold = true;
		///
		///     // Iterate through the paragraphs in this document.
		///     foreach (Paragraph p in document.Paragraphs)
		///     {
		///         /* 
		///          * Replace all instances of the string "wrong" with the string "right" and ignore case.
		///          * Each inserted instance of "wrong" should use the Formatting newFormatting.
		///          * Only replace an instance of "wrong" if it is Size 10, Italic and Times New Roman.
		///          * SubsetMatch means that the formatting must contain all elements of the match formatting,
		///          * but it can also contain additional formatting for example Color, UnderlineStyle, etc.
		///          * ExactMatch means it must not contain additional formatting.
		///          */
		///         p.ReplaceText("wrong", "right", false, RegexOptions.IgnoreCase, newFormatting, matchFormatting, MatchFormattingOptions.SubsetMatch);
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		/// <seealso cref="Paragraph.RemoveText(int, bool)"/>
		/// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		/// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		/// <param name="newValue">A System.String to replace all occurrences of oldValue.</param>
		/// <param name="oldValue">A System.String to be replaced.</param>
		/// <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
		/// <param name="trackChanges">Track changes</param>
		/// <param name="newFormatting">The formatting to apply to the text being inserted.</param>
		/// <param name="matchFormatting">The formatting that the text must match in order to be replaced.</param>
		/// <param name="fo">How should formatting be matched?</param>
		/// <param name="escapeRegEx">True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.</param>
		/// <param name="useRegExSubstitutions">True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).</param>
		public void ReplaceText(string oldValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
		{
			string tText = Text;
			MatchCollection mc = Regex.Matches(tText, escapeRegEx ? Regex.Escape(oldValue) : oldValue, options);

			// Loop through the matches in reverse order
			foreach (Match m in mc.Cast<Match>().Reverse())
			{
				// Assume the formatting matches until proven otherwise.
				bool formattingMatch = true;

				// Does the user want to match formatting?
				if (matchFormatting != null)
				{
					// The number of characters processed so far
					int processed = 0;

					do
					{
						// Get the next run effected
						Run run = GetFirstRunEffectedByEdit(m.Index + processed);

						// Get this runs properties
						XElement rPr = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));

						if (rPr == null)
							rPr = new Formatting().Xml;

						/* 
                         * Make sure that every formatting element in f.xml is also in this run,
                         * if this is not true, then their formatting does not match.
                         */
						if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
						{
							formattingMatch = false;
							break;
						}

						// We have processed some characters, so update the counter.
						processed += run.Value.Length;

					} while (processed < m.Length);
				}

				// If the formatting matches, do the replace.
				if (formattingMatch)
				{
					string repl = newValue;
					//perform RegEx substitutions. Only named groups are not supported. Everything else is supported. However character escapes are not covered.
					if (useRegExSubstitutions && !String.IsNullOrEmpty(repl))
					{
						repl = repl.Replace("$&", m.Value);
						if (m.Groups.Count > 0)
						{
							int lastcap = 0;
							for (int k = 0; k < m.Groups.Count; k++)
							{
								var g = m.Groups[k];
								if ((g == null) || (g.Value == ""))
									continue;
								repl = repl.Replace("$" + k.ToString(), g.Value);
								lastcap = k;
								//cannot get named groups ATM
							}
							repl = repl.Replace("$+", m.Groups[lastcap].Value);
						}
						if (m.Index > 0)
						{
							repl = repl.Replace("$`", tText.Substring(0, m.Index));
						}
						if ((m.Index + m.Length) < tText.Length)
						{
							repl = repl.Replace("$'", tText.Substring(m.Index + m.Length));
						}
						repl = repl.Replace("$_", tText);
						repl = repl.Replace("$$", "$");
					}
					if (!String.IsNullOrEmpty(repl))
						InsertText(m.Index + m.Length, repl, trackChanges, newFormatting);
					if (m.Length > 0)
						RemoveText(m.Index, m.Length, trackChanges);
				}
			}
		}

		/// <summary>
		/// Find pattern regex must return a group match.
		/// </summary>
		/// <param name="findPattern">Regex pattern that must include one group match. ie (.*)</param>
		/// <param name="regexMatchHandler">A func that accepts the matching find grouping text and returns a replacement value</param>
		/// <param name="trackChanges"></param>
		/// <param name="options"></param>
		/// <param name="newFormatting"></param>
		/// <param name="matchFormatting"></param>
		/// <param name="fo"></param>
		public void ReplaceText(string findPattern, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
		{
			var matchCollection = Regex.Matches(Text, findPattern, options);

			// Loop through the matches in reverse order
			foreach (var match in matchCollection.Cast<Match>().Reverse())
			{
				// Assume the formatting matches until proven otherwise.
				bool formattingMatch = true;

				// Does the user want to match formatting?
				if (matchFormatting != null)
				{
					// The number of characters processed so far
					int processed = 0;

					do
					{
						// Get the next run effected
						Run run = GetFirstRunEffectedByEdit(match.Index + processed);

						// Get this runs properties
						XElement rPr = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));

						if (rPr == null)
							rPr = new Formatting().Xml;

						/* 
                         * Make sure that every formatting element in f.xml is also in this run,
                         * if this is not true, then their formatting does not match.
                         */
						if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
						{
							formattingMatch = false;
							break;
						}

						// We have processed some characters, so update the counter.
						processed += run.Value.Length;

					} while (processed < match.Length);
				}

				// If the formatting matches, do the replace.
				if (formattingMatch)
				{
					var newValue = regexMatchHandler.Invoke(match.Groups[1].Value);
					InsertText(match.Index + match.Value.Length, newValue, trackChanges, newFormatting);
					RemoveText(match.Index, match.Value.Length, trackChanges);
				}
			}
		}


		/// <summary>
		/// Find all instances of a string in this paragraph and return their indexes in a List.
		/// </summary>
		/// <param name="str">The string to find</param>
		/// <returns>A list of indexes.</returns>
		/// <example>
		/// Find all instances of Hello in this document and insert 'don't' in frount of them.
		/// <code>
		/// // Load a document
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Loop through the paragraphs in this document.
		///     foreach(Paragraph p in document.Paragraphs)
		///     {
		///         // Find all instances of 'go' in this paragraph.
		///         <![CDATA[ List<int> ]]> gos = document.FindAll("go");
		///
		///         /* 
		///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
		///          * An important trick here is to do the inserting in reverse document order. If you inserted 
		///          * in document order, every insert would shift the index of the remaining matches.
		///          */
		///         gos.Reverse();
		///         foreach (int index in gos)
		///         {
		///             p.InsertText(index, "don't ", false);
		///         }
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public List<int> FindAll(string str)
		{
			return FindAll(str, RegexOptions.None);
		}

		/// <summary>
		/// Find all instances of a string in this paragraph and return their indexes in a List.
		/// </summary>
		/// <param name="str">The string to find</param>
		/// <param name="options">The options to use when finding a string match.</param>
		/// <returns>A list of indexes.</returns>
		/// <example>
		/// Find all instances of Hello in this document and insert 'don't' in frount of them.
		/// <code>
		/// // Load a document
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Loop through the paragraphs in this document.
		///     foreach(Paragraph p in document.Paragraphs)
		///     {
		///         // Find all instances of 'go' in this paragraph (Ignore case).
		///         <![CDATA[ List<int> ]]> gos = document.FindAll("go", RegexOptions.IgnoreCase);
		///
		///         /* 
		///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
		///          * An important trick here is to do the inserting in reverse document order. If you inserted 
		///          * in document order, every insert would shift the index of the remaining matches.
		///          */
		///         gos.Reverse();
		///         foreach (int index in gos)
		///         {
		///             p.InsertText(index, "don't ", false);
		///         }
		///     }
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public List<int> FindAll(string str, RegexOptions options)
		{
			MatchCollection mc = Regex.Matches(this.Text, Regex.Escape(str), options);

			var query =
			(
				from m in mc.Cast<Match>()
				select m.Index
			).ToList();

			return query;
		}

		/// <summary>
		///  Find all unique instances of the given Regex Pattern
		/// </summary>
		/// <param name="str"></param>
		/// <param name="options"></param>
		/// <returns></returns>
		public List<string> FindAllByPattern(string str, RegexOptions options)
		{
			MatchCollection mc = Regex.Matches(this.Text, str, options);

			var query =
			(
				from m in mc.Cast<Match>()
				select m.Value
			).ToList();

			return query;
		}

		/// <summary>
		/// Insert a PageNumber place holder into a Paragraph.
		/// This place holder should only be inserted into a Header or Footer Paragraph.
		/// Word will not automatically update this field if it is inserted into a document level Paragraph.
		/// </summary>
		/// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		/// <param name="index">The text index to insert this PageNumber place holder at.</param>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add Headers to the document.
		///     document.AddHeaders();
		///
		///     // Get the default Header.
		///     Header header = document.Headers.odd;
		///
		///     // Insert a Paragraph into the Header.
		///     Paragraph p0 = header.InsertParagraph("Page ( of )");
		///
		///     // Insert place holders for PageNumber and PageCount into the Header.
		///     // Word will replace these with the correct value for each Page.
		///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
		///     p0.InsertPageCount(PageNumberFormat.normal, 11);
		///
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AppendPageCount"/>
		/// <seealso cref="AppendPageNumber"/>
		/// <seealso cref="InsertPageCount"/>
		public void InsertPageNumber(PageNumberFormat pnf, int index = 0)
		{
			XElement fldSimple = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));

			if (pnf == PageNumberFormat.normal)
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" PAGE   \* MERGEFORMAT "));
			else
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" PAGE  \* ROMAN  \* MERGEFORMAT "));

			XElement content = XElement.Parse
			(
			 @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
			);

			fldSimple.Add(content);

			if (index == 0)
				Xml.AddFirst(fldSimple);
			else
			{
				Run r = GetFirstRunEffectedByEdit(index, EditType.ins);
				XElement[] splitEdit = SplitEdit(r.Xml, index, EditType.ins);
				r.Xml.ReplaceWith
				(
					splitEdit[0],
					fldSimple,
					splitEdit[1]
				);
			}
		}

		/// <summary>
		/// Append a PageNumber place holder onto the end of a Paragraph.
		/// </summary>
		/// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add Headers to the document.
		///     document.AddHeaders();
		///
		///     // Get the default Header.
		///     Header header = document.Headers.odd;
		///
		///     // Insert a Paragraph into the Header.
		///     Paragraph p0 = header.InsertParagraph();
		///
		///     // Appemd place holders for PageNumber and PageCount into the Header.
		///     // Word will replace these with the correct value for each Page.
		///     p0.Append("Page (");
		///     p0.AppendPageNumber(PageNumberFormat.normal);
		///     p0.Append(" of ");
		///     p0.AppendPageCount(PageNumberFormat.normal);
		///     p0.Append(")");
		/// 
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AppendPageCount"/>
		/// <seealso cref="InsertPageNumber"/>
		/// <seealso cref="InsertPageCount"/>
		public void AppendPageNumber(PageNumberFormat pnf)
		{
			XElement fldSimple = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));

			if (pnf == PageNumberFormat.normal)
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" PAGE   \* MERGEFORMAT "));
			else
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" PAGE  \* ROMAN  \* MERGEFORMAT "));

			XElement content = XElement.Parse
			(
			 @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
			);

			fldSimple.Add(content);
			Xml.Add(fldSimple);
		}

		/// <summary>
		/// Insert a PageCount place holder into a Paragraph.
		/// This place holder should only be inserted into a Header or Footer Paragraph.
		/// Word will not automatically update this field if it is inserted into a document level Paragraph.
		/// </summary>
		/// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		/// <param name="index">The text index to insert this PageCount place holder at.</param>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add Headers to the document.
		///     document.AddHeaders();
		///
		///     // Get the default Header.
		///     Header header = document.Headers.odd;
		///
		///     // Insert a Paragraph into the Header.
		///     Paragraph p0 = header.InsertParagraph("Page ( of )");
		///
		///     // Insert place holders for PageNumber and PageCount into the Header.
		///     // Word will replace these with the correct value for each Page.
		///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
		///     p0.InsertPageCount(PageNumberFormat.normal, 11);
		///
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AppendPageCount"/>
		/// <seealso cref="AppendPageNumber"/>
		/// <seealso cref="InsertPageNumber"/>
		public void InsertPageCount(PageNumberFormat pnf, int index = 0)
		{
			XElement fldSimple = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));

			if (pnf == PageNumberFormat.normal)
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" NUMPAGES   \* MERGEFORMAT "));
			else
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));

			XElement content = XElement.Parse
			(
			 @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
			);

			fldSimple.Add(content);

			if (index == 0)
				Xml.AddFirst(fldSimple);
			else
			{
				Run r = GetFirstRunEffectedByEdit(index, EditType.ins);
				XElement[] splitEdit = SplitEdit(r.Xml, index, EditType.ins);
				r.Xml.ReplaceWith
				(
					splitEdit[0],
					fldSimple,
					splitEdit[1]
				);
			}
		}

		/// <summary>
		/// Append a PageCount place holder onto the end of a Paragraph.
		/// </summary>
		/// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Add Headers to the document.
		///     document.AddHeaders();
		///
		///     // Get the default Header.
		///     Header header = document.Headers.odd;
		///
		///     // Insert a Paragraph into the Header.
		///     Paragraph p0 = header.InsertParagraph();
		///
		///     // Appemd place holders for PageNumber and PageCount into the Header.
		///     // Word will replace these with the correct value for each Page.
		///     p0.Append("Page (");
		///     p0.AppendPageNumber(PageNumberFormat.normal);
		///     p0.Append(" of ");
		///     p0.AppendPageCount(PageNumberFormat.normal);
		///     p0.Append(")");
		/// 
		///     // Save the document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		/// <seealso cref="AppendPageNumber"/>
		/// <seealso cref="InsertPageNumber"/>
		/// <seealso cref="InsertPageCount"/>
		public void AppendPageCount(PageNumberFormat pnf)
		{
			XElement fldSimple = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));

			if (pnf == PageNumberFormat.normal)
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" NUMPAGES   \* MERGEFORMAT "));
			else
				fldSimple.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));

			XElement content = XElement.Parse
			(
			 @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof /> 
                   </w:rPr>
                   <w:t>1</w:t> 
               </w:r>"
			);

			fldSimple.Add(content);
			Xml.Add(fldSimple);
		}

		/// <summary></summary>
		public float LineSpacingBefore
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));

				if (spacing != null)
				{
					XAttribute line = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName));
					if (line != null)
					{
						float f;

						if (float.TryParse(line.Value, out f))
							return f / 20.0f;
					}
				}

				return 0.0f;
			}

			set
			{
				SpacingBefore(value);
			}
		}

		/// <summary></summary>
		public float LineSpacingAfter
		{
			get
			{
				XElement pPr = GetOrCreate_pPr();
				XElement spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));

				if (spacing != null)
				{
					XAttribute line = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName));
					if (line != null)
					{
						float f;

						if (float.TryParse(line.Value, out f))
							return f / 20.0f;
					}
				}

				return 10.0f;
			}


			set
			{
				SpacingAfter(value);
			}
		}
	}

	/// <summary></summary>
	public class Run : DocXElement
	{

		// A lookup for the text elements in this paragraph
		Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

		private int startIndex;

		private int endIndex;

		private string text;

		/// <summary>
		/// Gets the start index of this Text (text length before this text)
		/// </summary>
		public int StartIndex { get { return startIndex; } }

		/// <summary>
		/// Gets the end index of this Text (text length before this text + this texts length)
		/// </summary>
		public int EndIndex { get { return endIndex; } }

		/// <summary>
		/// The text value of this text element
		/// </summary>
		internal string Value { set { text = value; } get { return text; } }

		internal Run(DocX document, XElement xml, int startIndex)
			: base(document, xml)
		{
			this.startIndex = startIndex;

			// Get the text elements in this run
			IEnumerable<XElement> texts = xml.Descendants();

			int start = startIndex;

			// Loop through each text in this run
			foreach (XElement te in texts)
			{
				switch (te.Name.LocalName)
				{
					case "tab":
						{
							textLookup.Add(start + 1, new Text(Document, te, start));
							text += "\t";
							start++;
							break;
						}
					case "br":
						{
							textLookup.Add(start + 1, new Text(Document, te, start));
							text += "\n";
							start++;
							break;
						}
					case "t": goto case "delText";
					case "delText":
						{
							// Only add strings which are not empty
							if (te.Value.Length > 0)
							{
								textLookup.Add(start + te.Value.Length, new Text(Document, te, start));
								text += te.Value;
								start += te.Value.Length;
							}
							break;
						}
					default: break;
				}
			}

			endIndex = start;
		}

		static internal XElement[] SplitRun(Run r, int index, EditType type = EditType.ins)
		{
			index = index - r.StartIndex;

			Text t = r.GetFirstTextEffectedByEdit(index, type);
			XElement[] splitText = Text.SplitText(t, index);

			XElement splitLeft = new XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), t.Xml.ElementsBeforeSelf().Where(n => n.Name.LocalName != "rPr"), splitText[0]);
			if (Paragraph.GetElementTextLength(splitLeft) == 0)
				splitLeft = null;

			XElement splitRight = new XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), splitText[1], t.Xml.ElementsAfterSelf().Where(n => n.Name.LocalName != "rPr"));
			if (Paragraph.GetElementTextLength(splitRight) == 0)
				splitRight = null;

			return
			(
				new XElement[]
				{
					splitLeft,
					splitRight
				}
			);
		}

		internal Text GetFirstTextEffectedByEdit(int index, EditType type = EditType.ins)
		{
			// Make sure we are looking within an acceptable index range.
			if (index < 0 || index > HelperFunctions.GetText(Xml).Length)
				throw new ArgumentOutOfRangeException();

			// Need some memory that can be updated by the recursive search for the XElement to Split.
			int count = 0;
			Text theOne = null;

			GetFirstTextEffectedByEditRecursive(Xml, index, ref count, ref theOne, type);

			return theOne;
		}

		internal void GetFirstTextEffectedByEditRecursive(XElement Xml, int index, ref int count, ref Text theOne, EditType type = EditType.ins)
		{
			count += HelperFunctions.GetSize(Xml);
			if (count > 0 && ((type == EditType.del && count > index) || (type == EditType.ins && count >= index)))
			{
				theOne = new Text(Document, Xml, count - HelperFunctions.GetSize(Xml));
				return;
			}

			if (Xml.HasElements)
				foreach (XElement e in Xml.Elements())
					if (theOne == null)
						GetFirstTextEffectedByEditRecursive(e, index, ref count, ref theOne);
		}
	}

	/// <summary>Represents a Picture in this document, a Picture is a customized view of an Image</summary>
	public class Picture : DocXElement
	{
		private const int EmusInPixel = 9525;

		internal Dictionary<PackagePart, PackageRelationship> picture_rels;

		internal Image img;
		private string id;
		private string name;
		private string descr;
		private int cx, cy;
		//private string fileName;
		private uint rotation;
		private bool hFlip, vFlip;
		private object pictureShape;
		private XElement xfrm;
		private XElement prstGeom;

		/// <summary>
		/// Remove this Picture from this document.
		/// </summary>
		public void Remove()
		{
			Xml.Remove();
		}

		/// <summary>
		/// Wraps an XElement as an Image
		/// </summary>
		/// <param name="document"></param>
		/// <param name="i">The XElement i to wrap</param>
		/// <param name="img"></param>
		internal Picture(DocX document, XElement i, Image img) : base(document, i)
		{
			picture_rels = new Dictionary<PackagePart, PackageRelationship>();

			this.img = img;

			this.id =
			(
				from e in Xml.Descendants()
				where e.Name.LocalName.Equals("blip")
				select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
			).SingleOrDefault();

			if (this.id == null)
			{
				this.id =
				(
					from e in Xml.Descendants()
					where e.Name.LocalName.Equals("imagedata")
					select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value
				).SingleOrDefault();
			}

			this.name =
			(
				from e in Xml.Descendants()
				let a = e.Attribute(XName.Get("name"))
				where (a != null)
				select a.Value
			).FirstOrDefault();

			if (this.name == null)
			{
				this.name =
				(
					from e in Xml.Descendants()
					let a = e.Attribute(XName.Get("title"))
					where (a != null)
					select a.Value
				).FirstOrDefault();
			}

			this.descr =
			(
				from e in Xml.Descendants()
				let a = e.Attribute(XName.Get("descr"))
				where (a != null)
				select a.Value
			).FirstOrDefault();

			this.cx =
			(
				from e in Xml.Descendants()
				let a = e.Attribute(XName.Get("cx"))
				where (a != null)
				select int.Parse(a.Value)
			).FirstOrDefault();

			if (this.cx == 0)
			{
				XAttribute style =
				(
					from e in Xml.Descendants()
					let a = e.Attribute(XName.Get("style"))
					where (a != null)
					select a
				).FirstOrDefault();

				string fromWidth = style.Value.Substring(style.Value.IndexOf("width:") + 6);
				var widthInt = ((double.Parse((fromWidth.Substring(0, fromWidth.IndexOf("pt"))).Replace(".", ","))) / 72.0) * 914400;
				cx = System.Convert.ToInt32(widthInt);
			}

			this.cy =
			(
				from e in Xml.Descendants()
				let a = e.Attribute(XName.Get("cy"))
				where (a != null)
				select int.Parse(a.Value)
			).FirstOrDefault();

			if (this.cy == 0)
			{
				XAttribute style =
				(
					from e in Xml.Descendants()
					let a = e.Attribute(XName.Get("style"))
					where (a != null)
					select a
				).FirstOrDefault();

				string fromHeight = style.Value.Substring(style.Value.IndexOf("height:") + 7);
				var heightInt = ((double.Parse((fromHeight.Substring(0, fromHeight.IndexOf("pt"))).Replace(".", ","))) / 72.0) * 914400;
				cy = System.Convert.ToInt32(heightInt);
			}

			this.xfrm =
			(
				from d in Xml.Descendants()
				where d.Name.LocalName.Equals("xfrm")
				select d
			).SingleOrDefault();

			this.prstGeom =
			(
				from d in Xml.Descendants()
				where d.Name.LocalName.Equals("prstGeom")
				select d
			).SingleOrDefault();

			if (xfrm != null)
				this.rotation = xfrm.Attribute(XName.Get("rot")) == null ? 0 : uint.Parse(xfrm.Attribute(XName.Get("rot")).Value);
		}

		private void SetPictureShape(object shape)
		{
			this.pictureShape = shape;

			XAttribute prst = prstGeom.Attribute(XName.Get("prst"));
			if (prst == null)
				prstGeom.Add(new XAttribute(XName.Get("prst"), "rectangle"));

			prstGeom.Attribute(XName.Get("prst")).Value = shape.ToString();
		}

		/// <summary>
		/// Set the shape of this Picture to one in the BasicShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the BasicShapes enumeration.</param>
		public void SetPictureShape(BasicShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the RectangleShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the RectangleShapes enumeration.</param>
		public void SetPictureShape(RectangleShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the BlockArrowShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
		public void SetPictureShape(BlockArrowShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the EquationShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the EquationShapes enumeration.</param>
		public void SetPictureShape(EquationShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the FlowchartShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the FlowchartShapes enumeration.</param>
		public void SetPictureShape(FlowchartShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
		public void SetPictureShape(StarAndBannerShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// Set the shape of this Picture to one in the CalloutShapes enumeration.
		/// </summary>
		/// <param name="shape">A shape from the CalloutShapes enumeration.</param>
		public void SetPictureShape(CalloutShapes shape)
		{
			SetPictureShape((object)shape);
		}

		/// <summary>
		/// A unique id that identifies an Image embedded in this document.
		/// </summary>
		public string Id
		{
			get { return id; }
		}

		/// <summary>
		/// Flip this Picture Horizontally.
		/// </summary>
		public bool FlipHorizontal
		{
			get { return hFlip; }

			set
			{
				hFlip = value;

				XAttribute flipH = xfrm.Attribute(XName.Get("flipH"));
				if (flipH == null)
					xfrm.Add(new XAttribute(XName.Get("flipH"), "0"));

				xfrm.Attribute(XName.Get("flipH")).Value = hFlip ? "1" : "0";
			}
		}

		/// <summary>
		/// Flip this Picture Vertically.
		/// </summary>
		public bool FlipVertical
		{
			get { return vFlip; }

			set
			{
				vFlip = value;

				XAttribute flipV = xfrm.Attribute(XName.Get("flipV"));
				if (flipV == null)
					xfrm.Add(new XAttribute(XName.Get("flipV"), "0"));

				xfrm.Attribute(XName.Get("flipV")).Value = vFlip ? "1" : "0";
			}
		}

		/// <summary>
		/// The rotation in degrees of this image, actual value = value % 360
		/// </summary>
		public uint Rotation
		{
			get { return rotation / 60000; }

			set
			{
				rotation = (value % 360) * 60000;
				XElement xfrm =
					(from d in Xml.Descendants()
					 where d.Name.LocalName.Equals("xfrm")
					 select d).Single();

				XAttribute rot = xfrm.Attribute(XName.Get("rot"));
				if (rot == null)
					xfrm.Add(new XAttribute(XName.Get("rot"), 0));

				xfrm.Attribute(XName.Get("rot")).Value = rotation.ToString();
			}
		}

		/// <summary>
		/// Gets or sets the name of this Image.
		/// </summary>
		public string Name
		{
			get { return name; }

			set
			{
				name = value;

				foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("name")))
					a.Value = name;
			}
		}

		/// <summary>
		/// Gets or sets the description for this Image.
		/// </summary>
		public string Description
		{
			get { return descr; }

			set
			{
				descr = value;

				foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("descr")))
					a.Value = descr;
			}
		}

		///<summary>
		/// Returns the name of the image file for the picture.
		///</summary>
		public string FileName
		{
			get
			{
				return img.FileName;
			}
		}

		/// <summary>
		/// Get or sets the Width of this Image.
		/// </summary>
		public int Width
		{
			get { return cx / EmusInPixel; }

			set
			{
				cx = value;

				foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("cx")))
					a.Value = (cx * EmusInPixel).ToString();
			}
		}

		/// <summary>
		/// Get or sets the height of this Image.
		/// </summary>
		public int Height
		{
			get { return cy / EmusInPixel; }

			set
			{
				cy = value;

				foreach (XAttribute a in Xml.Descendants().Attributes(XName.Get("cy")))
					a.Value = (cy * EmusInPixel).ToString();
			}
		}

		//public void Delete()
		//{
		//    // Remove xml
		//    i.Remove();

		//    // Rebuild the image collection for this paragraph
		//    // Requires that every Image have a link to its paragraph

		//}
	}

	/// <summary></summary>
	public class Section : Container
	{

		/// <summary></summary>
		public SectionBreakType SectionBreakType;

		internal Section(DocX document, XElement xml) : base(document, xml)
		{
		}

		/// <summary></summary>
		public List<Paragraph> SectionParagraphs { get; set; }

	}

	/// <summary>Represents a Table in a document</summary>
	public class Table : InsertBeforeOrAfter
	{

		private Alignment alignment;

		private AutoFit autofit;

		private float[] ColumnWidthsValue;

		/// <summary>
		/// Merge cells in given column starting with startRow and ending with endRow.
		/// </summary>
		/// <remarks>
		/// Added by arudoy patch: 11608
		/// </remarks>
		public void MergeCellsInColumn(int columnIndex, int startRow, int endRow)
		{
			// Check for valid start and end indexes.
			if (columnIndex < 0 || columnIndex >= ColumnCount)
				throw new IndexOutOfRangeException();

			if (startRow < 0 || endRow <= startRow || endRow >= Rows.Count)
				throw new IndexOutOfRangeException();
			// Foreach each Cell between startIndex and endIndex inclusive.
			foreach (Row row in Rows.Where((z, i) => i > startRow && i <= endRow))
			{
				Cell c = row.Cells[columnIndex];
				XElement tcPr = c.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					c.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = c.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				XElement vMerge = tcPr.Element(XName.Get("vMerge", DocX.w.NamespaceName));
				if (vMerge == null)
				{
					tcPr.SetElementValue(XName.Get("vMerge", DocX.w.NamespaceName), string.Empty);
					vMerge = tcPr.Element(XName.Get("vMerge", DocX.w.NamespaceName));
				}
			}

			/* 
             * Get the tcPr (table cell properties) element for the first cell in this merge,
            * null will be returned if no such element exists.
             */
			XElement start_tcPr = null;
			if (columnIndex > Rows[startRow].Cells.Count)
				start_tcPr = Rows[startRow].Cells[Rows[startRow].Cells.Count - 1].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			else
				start_tcPr = Rows[startRow].Cells[columnIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (start_tcPr == null)
			{
				Rows[startRow].Cells[columnIndex].Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				start_tcPr = Rows[startRow].Cells[columnIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}

			/* 
              * Get the gridSpan element of this row,
              * null will be returned if no such element exists.
              */
			XElement start_vMerge = start_tcPr.Element(XName.Get("vMerge", DocX.w.NamespaceName));
			if (start_vMerge == null)
			{
				start_tcPr.SetElementValue(XName.Get("vMerge", DocX.w.NamespaceName), string.Empty);
				start_vMerge = start_tcPr.Element(XName.Get("vMerge", DocX.w.NamespaceName));
			}

			start_vMerge.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "restart");
		}

		/// <summary>
		/// Returns a list of all Paragraphs inside this container.
		/// </summary>
		/// 
		public virtual List<Paragraph> Paragraphs
		{
			get
			{
				List<Paragraph> paragraphs = new List<Paragraph>();

				foreach (Row r in Rows)
					paragraphs.AddRange(r.Paragraphs);

				return paragraphs;
			}
		}

		/// <summary>
		/// Returns a list of all Pictures in a Table.
		/// </summary>
		/// <example>
		/// Returns a list of all Pictures in a Table.
		/// <code>
		/// <![CDATA[
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Get the first Table in a document.
		///     Table t = document.Tables[0];
		///
		///     // Get all of the Pictures in this Table.
		///     List<Picture> pictures = t.Pictures;
		///
		///     // Save this document.
		///     document.Save();
		/// }
		/// ]]>
		/// </code>
		/// </example>
		public List<Picture> Pictures
		{
			get
			{
				List<Picture> pictures = new List<Picture>();

				foreach (Row r in Rows)
					pictures.AddRange(r.Pictures);

				return pictures;
			}
		}

		/// <summary>
		/// Set the direction of all content in this Table.
		/// </summary>
		/// <param name="direction">(Left to Right) or (Right to Left)</param>
		/// <example>
		/// Set the content direction for all content in a table to RightToLeft.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Get the first table in a document.
		///     Table table = document.Tables[0];
		///
		///     // Set the content direction for all content in this table to RightToLeft.
		///     table.SetDirection(Direction.RightToLeft);
		///    
		///     // Save all changes made to this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public void SetDirection(Direction direction)
		{
			XElement tblPr = GetOrCreate_tblPr();
			tblPr.Add(new XElement(DocX.w + "bidiVisual"));

			foreach (Row r in Rows)
				r.SetDirection(direction);
		}

		/// <summary>
		/// Get all of the Hyperlinks in this Table.
		/// </summary>
		/// <example>
		/// Get all of the Hyperlinks in this Table.
		/// <code>
		/// // Create a document.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Get the first Table in this document.
		///     Table t = document.Tables[0];
		///
		///     // Get a list of all Hyperlinks in this Table.
		///     List&lt;Hyperlink&gt; hyperlinks = t.Hyperlinks;
		///
		///     // Save this document.
		///     document.Save();
		/// }
		/// </code>
		/// </example>
		public List<Hyperlink> Hyperlinks
		{
			get
			{
				List<Hyperlink> hyperlinks = new List<Hyperlink>();

				foreach (Row r in Rows)
					hyperlinks.AddRange(r.Hyperlinks);

				return hyperlinks;
			}
		}

		/// <summary></summary>
		/// <param name="widths"></param>
		public void SetWidths(float[] widths)
		{
			this.ColumnWidthsValue = widths;
			//set widths for existing rows
			foreach (var r in Rows)
			{
				for (var c = 0; c < widths.Length; c++)
				{
					if (r.Cells.Count > c)
						r.Cells[c].Width = widths[c];
				}

			}
		}

		/// <summary> 
		/// Set Table column width by prescribing percent 
		/// </summary> 
		/// <param name="widthsPercentage">column width % list</param> 
		/// <param name="totalWidth">Total table width. Will be calculated if null sent.</param>
		public void SetWidthsPercentage(float[] widthsPercentage, float? totalWidth)
		{
			if (totalWidth == null) totalWidth = this.Document.PageWidth - this.Document.MarginLeft - this.Document.MarginRight; // calculate total table width 
			List<float> widths = new List<float>(widthsPercentage.Length); // empty list, will hold actual width 
			widthsPercentage.ToList().ForEach(pWidth => { widths.Add(pWidth * totalWidth.Value / 100); }); // convert percentage to actual width for all values in array 
			SetWidths(widths.ToArray()); // set actual column width
		}


		/// <summary>
		/// If the tblPr element doesent exist it is created, either way it is returned by this function.
		/// </summary>
		/// <returns>The tblPr element for this Table.</returns>
		internal XElement GetOrCreate_tblPr()
		{
			// Get the element.
			XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));

			// If it dosen't exist, create it.
			if (tblPr == null)
			{
				Xml.AddFirst(new XElement(XName.Get("tblPr", DocX.w.NamespaceName)));
				tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			}

			// Return the pPr element for this Paragraph.
			return tblPr;
		}

		/// <summary>Set the specified cell margin for the table-level</summary>
		/// <param name="type">The side of the cell margin</param>
		/// <param name="margin">The value for the specified cell margin</param>
		/// <remarks>More information can be found here http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellmargindefault.aspx</remarks>
		public void SetTableCellMargin(TableCellMarginType type, double margin)
		{
			XElement tblPr = GetOrCreate_tblPr();

			// find (or create) the element with the cell margins 
			XElement tblCellMar = tblPr.Element(XName.Get("tblCellMar", DocX.w.NamespaceName));
			if (tblCellMar == null)
			{
				tblPr.AddFirst(new XElement(XName.Get("tblCellMar", DocX.w.NamespaceName)));
				tblCellMar = tblPr.Element(XName.Get("tblCellMar", DocX.w.NamespaceName));
			}

			// find (or create) the element with cell margin for the specified side
			XElement tblMargin = tblCellMar.Element(XName.Get(type.ToString(), DocX.w.NamespaceName));
			if (tblMargin == null)
			{
				tblCellMar.AddFirst(new XElement(XName.Get(type.ToString(), DocX.w.NamespaceName)));
				tblMargin = tblCellMar.Element(XName.Get(type.ToString(), DocX.w.NamespaceName));
			}

			tblMargin.RemoveAttributes();
			// set the value for the cell margin
			tblMargin.Add(new XAttribute(XName.Get("w", DocX.w.NamespaceName), margin));
			// set the side of cell margin
			tblMargin.Add(new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"));
		}

		/// <summary>
		/// Gets the column width for a given column index.
		/// </summary>
		/// <param name="index"></param>
		public Double GetColumnWidth(Int32 index)
		{
			List<Double> widths = ColumnWidths;
			if (widths == null || index > widths.Count - 1) return Double.NaN;

			return widths[index];
		}

		/// <summary>
		/// Sets the column width for the given index.
		/// </summary>
		/// <param name="index">Column index</param>
		/// <param name="width">Colum width</param>
		public void SetColumnWidth(Int32 index, Double width)
		{
			List<Double> widths = ColumnWidths;
			if (widths == null || index > widths.Count - 1)
			{
				if (Rows.Count == 0) throw new Exception("There is at least one row required to detect the existing columns.");
				// use width of last row cells
				// may not work for merged cell! 
				widths = new List<Double>();
				foreach (Cell c in Rows[Rows.Count - 1].Cells)
				{
					widths.Add(c.Width);
				}
			}

			// check if index is matching table columns
			if (index > widths.Count - 1) throw new Exception("The index is greather than the available table columns.");

			// get the table grid props
			XElement grid = Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName));
			// if null; append a new grid below tblPr
			if (grid == null)
			{
				XElement tblPr = GetOrCreate_tblPr();
				tblPr.AddAfterSelf(new XElement(XName.Get("tblGrid", DocX.w.NamespaceName)));
				grid = Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName));
			}

			// remove all existing values
			grid.RemoveAll();

			// append new column widths
			XElement gridCol = null;
			Int32 i = 0;
			Double value = width;
			Double total = 0;
			foreach (var w in widths)
			{
				value = w;
				if (i == index) value = width;
				gridCol = new XElement(XName.Get("gridCol", DocX.w.NamespaceName),
						  new XAttribute(XName.Get("w", DocX.w.NamespaceName), value));
				grid.Add(gridCol);
				i += 1;
				total += value;
			}

			// remove cell widths
			foreach (Row r in Rows)
				foreach (Cell c in r.Cells)
					c.Width = -1;

			// set fitting to fixed; this will add/set additional table properties
			this.AutoFit = AutoFit.Fixed;
		}


		/// <summary>
		/// Gets a list of all column widths for this table.
		/// </summary>
		public List<Double> ColumnWidths
		{
			get
			{
				List<Double> widths = new List<Double>();
				// get the table grid props
				XElement grid = Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName));
				if (grid == null) return null;

				// get col properties
				var cols = grid.Elements(XName.Get("gridCol", DocX.w.NamespaceName));
				if (cols == null) return null;

				String value = String.Empty;
				foreach (var col in cols)
				{
					value = col.GetAttribute(XName.Get("w", DocX.w.NamespaceName));
					widths.Add(Convert.ToDouble(value));
				}
				return widths;
			}
		}


		/// <summary>
		/// Returns the number of rows in this table.
		/// </summary>
		public Int32 RowCount
		{
			get
			{
				return Xml.Elements(XName.Get("tr", DocX.w.NamespaceName)).Count();
			}
		}

		private int _cachedColCount = -1;
		/// <summary>
		/// Returns the number of columns in this table.
		/// </summary>
		public Int32 ColumnCount
		{
			get
			{
				if (RowCount == 0)
					return 0;
				if (_cachedColCount == -1)
					_cachedColCount = Rows.First().ColumnCount;
				return _cachedColCount;
			}
		}

		/// <summary>
		/// Returns a list of rows in this table.
		/// </summary>
		public List<Row> Rows
		{
			get
			{
				List<Row> rows =
				(
					from r in Xml.Elements(XName.Get("tr", DocX.w.NamespaceName))
					select new Row(this, Document, r)
				).ToList();

				return rows;
			}
		}

		private TableDesign design;


		internal Table(DocX document, XElement xml)
			: base(document, xml)
		{
			autofit = AutoFit.ColumnWidth;
			this.Xml = xml;
			this.mainPart = document.mainPart;

			XElement properties = xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));

			XElement style = properties.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
			if (style != null)
			{
				XAttribute val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));

				if (val != null)
				{
					try
					{
						design = (TableDesign)Enum.Parse(typeof(TableDesign), val.Value.Replace("-", string.Empty));
					}

					catch (Exception)
					{
						design = TableDesign.Custom;
					}
				}
				else
					design = TableDesign.None;
			}

			else
				design = TableDesign.None;

			XElement tableLook = properties.Element(XName.Get("tblLook", DocX.w.NamespaceName));
			if (tableLook != null)
			{
				TableLook = new TableLook();
				TableLook.FirstRow = tableLook.GetAttribute(XName.Get("firstRow", DocX.w.NamespaceName)) == "1";
				TableLook.LastRow = tableLook.GetAttribute(XName.Get("lastRow", DocX.w.NamespaceName)) == "1";
				TableLook.FirstColumn = tableLook.GetAttribute(XName.Get("firstColumn", DocX.w.NamespaceName)) == "1";
				TableLook.LastColumn = tableLook.GetAttribute(XName.Get("lastColumn", DocX.w.NamespaceName)) == "1";
				TableLook.NoHorizontalBanding = tableLook.GetAttribute(XName.Get("noHBand", DocX.w.NamespaceName)) == "1";
				TableLook.NoVerticalBanding = tableLook.GetAttribute(XName.Get("noVBand", DocX.w.NamespaceName)) == "1";
			}

		}
		/// <summary>
		/// Extra property for Custom Table Style provided by carpfisher - Thanks
		/// </summary>
		private string _customTableDesignName;
		/// <summary>
		/// Extra property for Custom Table Style provided by carpfisher - Thanks
		/// </summary>
		public string CustomTableDesignName
		{
			set
			{
				_customTableDesignName = value;
				this.Design = TableDesign.Custom;
			}

			get
			{
				return _customTableDesignName;
			}
		}

		/// <summary>
		/// String containing the Table Caption value (the table's Alternate Text Title)
		/// </summary>
		private string _tableCaption;
		/// <summary>
		/// Gets or Sets the value of the Table Caption (Alternate Text Title) of this table. 
		/// </summary>
		public string TableCaption
		{
			set
			{
				XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (tblPr != null)
				{
					XElement tblCaption =
						tblPr.Descendants(XName.Get("tblCaption", DocX.w.NamespaceName)).FirstOrDefault();

					if (tblCaption != null)
						tblCaption.Remove();

					tblCaption = new XElement(XName.Get("tblCaption", DocX.w.NamespaceName),
						new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
					tblPr.Add(tblCaption);
				}
			}

			get
			{
				XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (tblPr != null)
				{
					XElement caption = tblPr.Element(XName.Get("tblCaption", DocX.w.NamespaceName));
					if (caption != null)
					{
						_tableCaption = caption.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
					}
				}
				return _tableCaption;
			}
		}

		/// <summary>
		/// String containing the Table Description (the table's Alternate Text Description).
		/// </summary>
		private string _tableDescription;
		/// <summary>
		/// Gets or Sets the value of the Table Description (Alternate Text Description) of this table. 
		/// </summary>
		public string TableDescription
		{
			set
			{
				XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (tblPr != null)
				{
					XElement tblDescription =
						tblPr.Descendants(XName.Get("tblDescription", DocX.w.NamespaceName)).FirstOrDefault();

					if (tblDescription != null)
						tblDescription.Remove();

					tblDescription = new XElement(XName.Get("tblDescription", DocX.w.NamespaceName),
					   new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
					tblPr.Add(tblDescription);
				}
			}

			get
			{
				XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (tblPr != null)
				{
					XElement caption = tblPr.Element(XName.Get("tblDescription", DocX.w.NamespaceName));
					if (caption != null)
					{
						_tableDescription = caption.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
					}
				}
				return _tableDescription;
			}
		}

		/// <summary></summary>
		public TableLook TableLook
		{
			get;
			set;
		}

		/// <summary></summary>
		public Alignment Alignment
		{
			get
			{
				return alignment;
			}
			set
			{
				string alignmentString = string.Empty;
				switch (value)
				{
					case Alignment.left:
						{
							alignmentString = "left";
							break;
						}
					case Alignment.both:
						{
							alignmentString = "both";
							break;
						}
					case Alignment.right:
						{
							alignmentString = "right";
							break;
						}
					case Alignment.center:
						{
							alignmentString = "center";
							break;
						}
				}
				XElement tblPr = Xml.Descendants(XName.Get("tblPr", DocX.w.NamespaceName)).First();
				XElement jc = tblPr.Descendants(XName.Get("jc", DocX.w.NamespaceName)).FirstOrDefault();
				if (jc != null) jc.Remove();
				jc = new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignmentString));
				tblPr.Add(jc);
				alignment = value;
			}
		}

		/// <summary>
		/// Auto size this table according to some rule.
		/// </summary>
		/// <remarks>Added by Roger Saele, April 2012. Thank you for your contribution Roger.</remarks>
		public AutoFit AutoFit
		{
			get
			{
				return autofit;
			}
			set
			{
				string tableAttributeValue = string.Empty;
				string columnAttributeValue = string.Empty;
				switch (value)
				{
					case AutoFit.ColumnWidth:
						{
							tableAttributeValue = "auto";
							columnAttributeValue = "dxa";

							// Disable "Automatically resize to fit contents" option
							XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
							if (tblPr != null)
							{
								XElement layout = tblPr.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
								if (layout == null)
								{
									tblPr.Add(new XElement(XName.Get("tblLayout", DocX.w.NamespaceName)));
									layout = tblPr.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
								}

								XAttribute type = layout.Attribute(XName.Get("type", DocX.w.NamespaceName));
								if (type == null)
								{
									layout.Add(new XAttribute(XName.Get("type", DocX.w.NamespaceName), String.Empty));
									type = layout.Attribute(XName.Get("type", DocX.w.NamespaceName));
								}

								type.Value = "fixed";
							}

							break;
						}

					case AutoFit.Contents:
						{
							tableAttributeValue = columnAttributeValue = "auto";
							break;
						}

					case AutoFit.Window:
						{
							tableAttributeValue = columnAttributeValue = "pct";
							break;
						}

					case AutoFit.Fixed:
						// DL added - 20150816:
						// Set fixed width for the whole table; columns width is definied in the node: tblGrid
						{
							tableAttributeValue = columnAttributeValue = "dxa";

							XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
							XElement tblLayout = tblPr.Element(XName.Get("tblLayout", DocX.w.NamespaceName));

							if (tblLayout == null)
							{
								XElement tmp = tblPr.Element(XName.Get("tblInd", DocX.w.NamespaceName));
								if (tmp == null)
								{
									tmp = tblPr.Element(XName.Get("tblW", DocX.w.NamespaceName));
								}

								tmp.AddAfterSelf(new XElement(XName.Get("tblLayout", DocX.w.NamespaceName)));
								tmp = tblPr.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
								tmp.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "fixed");

								tmp = tblPr.Element(XName.Get("tblW", DocX.w.NamespaceName));
								Double i = 0;
								foreach (Double w in ColumnWidths)
									i += w;

								tmp.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), i.ToString());


								break;

							}
							else
							{
								var qry = from d in Xml.Descendants()
										  let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
										  where (d.Name.LocalName == "tblLayout") && type != null
										  select type;

								foreach (XAttribute type in qry) type.Value = "fixed";
								XElement tmp = tblPr.Element(XName.Get("tblW", DocX.w.NamespaceName));
								Double i = 0;
								foreach (Double w in ColumnWidths) i += w;
								tmp.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), i.ToString());
								break;
							}
						}
				}
				// Set table attributes
				var query = from d in Xml.Descendants()
							let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
							where (d.Name.LocalName == "tblW") && type != null
							select type;

				foreach (XAttribute type in query)
					type.Value = tableAttributeValue;

				// Set column attributes
				query = from d in Xml.Descendants()
						let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
						where (d.Name.LocalName == "tcW") && type != null
						select type;

				foreach (XAttribute type in query)
					type.Value = columnAttributeValue;

				autofit = value;
			}
		}
		/// <summary>
		/// The design\style to apply to this table.
		/// 
		/// Patch1. Patch to code for Custom Table Style support by carpfisher
		/// </summary>
		/// <example>
		/// Example code for custom table style usage 
		/// 
		/// <code> 
		/// Novacode.DocX document = Novacode.DocX.Load(“DOC01.doc”); // load document with custom table style defined
		/// Novacode.Table t = document.AddTable(2, 2); // adds table 
		/// t.CustomTableDesignName = “MyStyle01”; // assigns Custom Table Design style to newly created table
		/// </code>
		/// </example>
		/// 
		/// 
		/// 
		public TableDesign Design
		{
			get { return design; }
			set
			{
				XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				XElement style = tblPr.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
				if (style == null)
				{
					tblPr.Add(new XElement(XName.Get("tblStyle", DocX.w.NamespaceName)));
					style = tblPr.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
				}

				XAttribute val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (val == null)
				{
					style.Add(new XAttribute(XName.Get("val", DocX.w.NamespaceName), ""));
					val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));
				}

				design = value;

				if (design == TableDesign.None)
				{
					if (style != null)
						style.Remove();
				}

				if (design == TableDesign.Custom)
				{
					if (string.IsNullOrEmpty(_customTableDesignName))
					{
						design = TableDesign.None;
						if (style != null)
							style.Remove();

					}
					else
					{
						val.Value = _customTableDesignName;
					}
				}
				else
				{

					switch (design)
					{
						case TableDesign.TableNormal:
							val.Value = "TableNormal";
							break;
						case TableDesign.TableGrid:
							val.Value = "TableGrid";
							break;
						case TableDesign.LightShading:
							val.Value = "LightShading";
							break;
						case TableDesign.LightShadingAccent1:
							val.Value = "LightShading-Accent1";
							break;
						case TableDesign.LightShadingAccent2:
							val.Value = "LightShading-Accent2";
							break;
						case TableDesign.LightShadingAccent3:
							val.Value = "LightShading-Accent3";
							break;
						case TableDesign.LightShadingAccent4:
							val.Value = "LightShading-Accent4";
							break;
						case TableDesign.LightShadingAccent5:
							val.Value = "LightShading-Accent5";
							break;
						case TableDesign.LightShadingAccent6:
							val.Value = "LightShading-Accent6";
							break;
						case TableDesign.LightList:
							val.Value = "LightList";
							break;
						case TableDesign.LightListAccent1:
							val.Value = "LightList-Accent1";
							break;
						case TableDesign.LightListAccent2:
							val.Value = "LightList-Accent2";
							break;
						case TableDesign.LightListAccent3:
							val.Value = "LightList-Accent3";
							break;
						case TableDesign.LightListAccent4:
							val.Value = "LightList-Accent4";
							break;
						case TableDesign.LightListAccent5:
							val.Value = "LightList-Accent5";
							break;
						case TableDesign.LightListAccent6:
							val.Value = "LightList-Accent6";
							break;
						case TableDesign.LightGrid:
							val.Value = "LightGrid";
							break;
						case TableDesign.LightGridAccent1:
							val.Value = "LightGrid-Accent1";
							break;
						case TableDesign.LightGridAccent2:
							val.Value = "LightGrid-Accent2";
							break;
						case TableDesign.LightGridAccent3:
							val.Value = "LightGrid-Accent3";
							break;
						case TableDesign.LightGridAccent4:
							val.Value = "LightGrid-Accent4";
							break;
						case TableDesign.LightGridAccent5:
							val.Value = "LightGrid-Accent5";
							break;
						case TableDesign.LightGridAccent6:
							val.Value = "LightGrid-Accent6";
							break;
						case TableDesign.MediumShading1:
							val.Value = "MediumShading1";
							break;
						case TableDesign.MediumShading1Accent1:
							val.Value = "MediumShading1-Accent1";
							break;
						case TableDesign.MediumShading1Accent2:
							val.Value = "MediumShading1-Accent2";
							break;
						case TableDesign.MediumShading1Accent3:
							val.Value = "MediumShading1-Accent3";
							break;
						case TableDesign.MediumShading1Accent4:
							val.Value = "MediumShading1-Accent4";
							break;
						case TableDesign.MediumShading1Accent5:
							val.Value = "MediumShading1-Accent5";
							break;
						case TableDesign.MediumShading1Accent6:
							val.Value = "MediumShading1-Accent6";
							break;
						case TableDesign.MediumShading2:
							val.Value = "MediumShading2";
							break;
						case TableDesign.MediumShading2Accent1:
							val.Value = "MediumShading2-Accent1";
							break;
						case TableDesign.MediumShading2Accent2:
							val.Value = "MediumShading2-Accent2";
							break;
						case TableDesign.MediumShading2Accent3:
							val.Value = "MediumShading2-Accent3";
							break;
						case TableDesign.MediumShading2Accent4:
							val.Value = "MediumShading2-Accent4";
							break;
						case TableDesign.MediumShading2Accent5:
							val.Value = "MediumShading2-Accent5";
							break;
						case TableDesign.MediumShading2Accent6:
							val.Value = "MediumShading2-Accent6";
							break;
						case TableDesign.MediumList1:
							val.Value = "MediumList1";
							break;
						case TableDesign.MediumList1Accent1:
							val.Value = "MediumList1-Accent1";
							break;
						case TableDesign.MediumList1Accent2:
							val.Value = "MediumList1-Accent2";
							break;
						case TableDesign.MediumList1Accent3:
							val.Value = "MediumList1-Accent3";
							break;
						case TableDesign.MediumList1Accent4:
							val.Value = "MediumList1-Accent4";
							break;
						case TableDesign.MediumList1Accent5:
							val.Value = "MediumList1-Accent5";
							break;
						case TableDesign.MediumList1Accent6:
							val.Value = "MediumList1-Accent6";
							break;
						case TableDesign.MediumList2:
							val.Value = "MediumList2";
							break;
						case TableDesign.MediumList2Accent1:
							val.Value = "MediumList2-Accent1";
							break;
						case TableDesign.MediumList2Accent2:
							val.Value = "MediumList2-Accent2";
							break;
						case TableDesign.MediumList2Accent3:
							val.Value = "MediumList2-Accent3";
							break;
						case TableDesign.MediumList2Accent4:
							val.Value = "MediumList2-Accent4";
							break;
						case TableDesign.MediumList2Accent5:
							val.Value = "MediumList2-Accent5";
							break;
						case TableDesign.MediumList2Accent6:
							val.Value = "MediumList2-Accent6";
							break;
						case TableDesign.MediumGrid1:
							val.Value = "MediumGrid1";
							break;
						case TableDesign.MediumGrid1Accent1:
							val.Value = "MediumGrid1-Accent1";
							break;
						case TableDesign.MediumGrid1Accent2:
							val.Value = "MediumGrid1-Accent2";
							break;
						case TableDesign.MediumGrid1Accent3:
							val.Value = "MediumGrid1-Accent3";
							break;
						case TableDesign.MediumGrid1Accent4:
							val.Value = "MediumGrid1-Accent4";
							break;
						case TableDesign.MediumGrid1Accent5:
							val.Value = "MediumGrid1-Accent5";
							break;
						case TableDesign.MediumGrid1Accent6:
							val.Value = "MediumGrid1-Accent6";
							break;
						case TableDesign.MediumGrid2:
							val.Value = "MediumGrid2";
							break;
						case TableDesign.MediumGrid2Accent1:
							val.Value = "MediumGrid2-Accent1";
							break;
						case TableDesign.MediumGrid2Accent2:
							val.Value = "MediumGrid2-Accent2";
							break;
						case TableDesign.MediumGrid2Accent3:
							val.Value = "MediumGrid2-Accent3";
							break;
						case TableDesign.MediumGrid2Accent4:
							val.Value = "MediumGrid2-Accent4";
							break;
						case TableDesign.MediumGrid2Accent5:
							val.Value = "MediumGrid2-Accent5";
							break;
						case TableDesign.MediumGrid2Accent6:
							val.Value = "MediumGrid2-Accent6";
							break;
						case TableDesign.MediumGrid3:
							val.Value = "MediumGrid3";
							break;
						case TableDesign.MediumGrid3Accent1:
							val.Value = "MediumGrid3-Accent1";
							break;
						case TableDesign.MediumGrid3Accent2:
							val.Value = "MediumGrid3-Accent2";
							break;
						case TableDesign.MediumGrid3Accent3:
							val.Value = "MediumGrid3-Accent3";
							break;
						case TableDesign.MediumGrid3Accent4:
							val.Value = "MediumGrid3-Accent4";
							break;
						case TableDesign.MediumGrid3Accent5:
							val.Value = "MediumGrid3-Accent5";
							break;
						case TableDesign.MediumGrid3Accent6:
							val.Value = "MediumGrid3-Accent6";
							break;

						case TableDesign.DarkList:
							val.Value = "DarkList";
							break;
						case TableDesign.DarkListAccent1:
							val.Value = "DarkList-Accent1";
							break;
						case TableDesign.DarkListAccent2:
							val.Value = "DarkList-Accent2";
							break;
						case TableDesign.DarkListAccent3:
							val.Value = "DarkList-Accent3";
							break;
						case TableDesign.DarkListAccent4:
							val.Value = "DarkList-Accent4";
							break;
						case TableDesign.DarkListAccent5:
							val.Value = "DarkList-Accent5";
							break;
						case TableDesign.DarkListAccent6:
							val.Value = "DarkList-Accent6";
							break;

						case TableDesign.ColorfulShading:
							val.Value = "ColorfulShading";
							break;
						case TableDesign.ColorfulShadingAccent1:
							val.Value = "ColorfulShading-Accent1";
							break;
						case TableDesign.ColorfulShadingAccent2:
							val.Value = "ColorfulShading-Accent2";
							break;
						case TableDesign.ColorfulShadingAccent3:
							val.Value = "ColorfulShading-Accent3";
							break;
						case TableDesign.ColorfulShadingAccent4:
							val.Value = "ColorfulShading-Accent4";
							break;
						case TableDesign.ColorfulShadingAccent5:
							val.Value = "ColorfulShading-Accent5";
							break;
						case TableDesign.ColorfulShadingAccent6:
							val.Value = "ColorfulShading-Accent6";
							break;

						case TableDesign.ColorfulList:
							val.Value = "ColorfulList";
							break;
						case TableDesign.ColorfulListAccent1:
							val.Value = "ColorfulList-Accent1";
							break;
						case TableDesign.ColorfulListAccent2:
							val.Value = "ColorfulList-Accent2";
							break;
						case TableDesign.ColorfulListAccent3:
							val.Value = "ColorfulList-Accent3";
							break;
						case TableDesign.ColorfulListAccent4:
							val.Value = "ColorfulList-Accent4";
							break;
						case TableDesign.ColorfulListAccent5:
							val.Value = "ColorfulList-Accent5";
							break;
						case TableDesign.ColorfulListAccent6:
							val.Value = "ColorfulList-Accent6";
							break;

						case TableDesign.ColorfulGrid:
							val.Value = "ColorfulGrid";
							break;
						case TableDesign.ColorfulGridAccent1:
							val.Value = "ColorfulGrid-Accent1";
							break;
						case TableDesign.ColorfulGridAccent2:
							val.Value = "ColorfulGrid-Accent2";
							break;
						case TableDesign.ColorfulGridAccent3:
							val.Value = "ColorfulGrid-Accent3";
							break;
						case TableDesign.ColorfulGridAccent4:
							val.Value = "ColorfulGrid-Accent4";
							break;
						case TableDesign.ColorfulGridAccent5:
							val.Value = "ColorfulGrid-Accent5";
							break;
						case TableDesign.ColorfulGridAccent6:
							val.Value = "ColorfulGrid-Accent6";
							break;

						default:
							break;
					}
				}
				if (Document.styles == null)
				{
					PackagePart word_styles = Document.package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
					using (TextReader tr = new StreamReader(word_styles.GetStream()))
						Document.styles = XDocument.Load(tr);
				}

				var tableStyle =
				(
					from e in Document.styles.Descendants()
					let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
					where (styleId != null && styleId.Value == val.Value)
					select e
				).FirstOrDefault();

				if (tableStyle == null)
				{
					XDocument external_style_doc = HelperFunctions.DecompressXMLResource("Novacode.Resources.styles.xml.gz");

					var styleElement =
					(
						from e in external_style_doc.Descendants()
						let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
						where (styleId != null && styleId.Value == val.Value)
						select e
					).First();

					Document.styles.Element(XName.Get("styles", DocX.w.NamespaceName)).Add(styleElement);
				}
			}
		}

		/// <summary>
		/// Returns the index of this Table.
		/// </summary>
		/// <example>
		/// Replace the first table in this document with a new Table.
		/// <code>
		/// // Load a document into memory.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Get the first Table in this document.
		///     Table t = document.Tables[0];
		///
		///     // Get the character index of Table t in this document.
		///     int index = t.Index;
		///
		///     // Remove Table t.
		///     t.Remove();
		///
		///     // Insert a new Table at the original index of Table t.
		///     Table newTable = document.InsertTable(index, 4, 4);
		///
		///     // Set the design of this new Table, so that we can see it.
		///     newTable.Design = TableDesign.LightShadingAccent1;
		///
		///     // Save all changes made to the document.
		///     document.Save();
		/// } // Release this document from memory.
		/// </code>
		/// </example>
		public int Index
		{
			get
			{
				int index = 0;
				IEnumerable<XElement> previous = Xml.ElementsBeforeSelf();

				foreach (XElement e in previous)
					index += Paragraph.GetElementTextLength(e);

				return index;
			}
		}

		/// <summary>
		/// Remove this Table from this document.
		/// </summary>
		/// <example>
		/// Remove the first Table from this document.
		/// <code>
		/// // Load a document into memory.
		/// using (DocX document = DocX.Load(@"Test.docx"))
		/// {
		///     // Get the first Table in this document.
		///     Table t = d.Tables[0];
		///        
		///     // Remove this Table.
		///     t.Remove();
		///
		///     // Save all changes made to the document.
		///     document.Save();
		/// } // Release this document from memory.
		/// </code>
		/// </example>
		public void Remove()
		{
			Xml.Remove();
		}

		/// <summary>
		/// Insert a row at the end of this table.
		/// </summary>
		/// <example>
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///        
		///     // Insert a new row at the end of this table.
		///     Row row = table.InsertRow();
		///
		///     // Loop through each cell in this new row.
		///     foreach (Cell c in row.Cells)
		///     {
		///         // Set the text of each new cell to "Hello".
		///         c.Paragraphs[0].InsertText("Hello", false);
		///     }
		///
		///     // Save the document to a new file.
		///     document.SaveAs(@"C:\Example\Test2.docx");
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <returns>A new row.</returns>
		public Row InsertRow()
		{
			return InsertRow(RowCount);
		}

		/// <summary>
		/// Insert a copy of a row at the end of this table.
		/// </summary>      
		/// <returns>A new row.</returns>
		public Row InsertRow(Row row)
		{
			return InsertRow(row, RowCount);
		}

		/// <summary>
		/// Insert a column to the right of a Table.
		/// </summary>
		/// <example>
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first Table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Insert a new column to this right of this table.
		///     table.InsertColumn();
		///
		///     // Set the new columns text to "Row no."
		///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
		///
		///     // Loop through each row in the table.
		///     for (int i = 1; i &lt; table.Rows.Count; i++)
		///     {
		///         // The current row.
		///         Row row = table.Rows[i];
		///
		///         // The cell in this row that belongs to the new column.
		///         Cell cell = row.Cells[table.ColumnCount - 1];
		///
		///         // The first Paragraph that this cell houses.
		///         Paragraph p = cell.Paragraphs[0];
		///
		///         // Insert this rows index.
		///         p.InsertText(i.ToString(), false);
		///     }
		///
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void InsertColumn()
		{
			InsertColumn(ColumnCount, true);
		}

		/// <summary>
		/// Remove the last row from this Table.
		/// </summary>
		/// <example>
		/// Remove the last row from a Table.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Remove the last row from this table.
		///     table.RemoveRow();
		///
		///     // Save the document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void RemoveRow()
		{
			RemoveRow(RowCount - 1);
		}

		/// <summary>
		/// Remove a row from this Table.
		/// </summary>
		/// <param name="index">The row to remove.</param>
		/// <example>
		/// Remove the first row from a Table.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Remove the first row from this table.
		///     table.RemoveRow(0);
		///
		///     // Save the document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void RemoveRow(int index)
		{
			if (index < 0 || index > RowCount - 1)
				throw new IndexOutOfRangeException();

			Rows[index].Xml.Remove();
			if (Rows.Count == 0)
				Remove();
		}

		/// <summary>
		/// Remove the last column for this Table.
		/// </summary>
		/// <example>
		/// Remove the last column from a Table.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Remove the last column from this table.
		///     table.RemoveColumn();
		///
		///     // Save the document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void RemoveColumn()
		{
			RemoveColumn(ColumnCount - 1);
		}

		/// <summary>
		/// Remove a column from this Table.
		/// </summary>
		/// <param name="index">The column to remove.</param>
		/// <example>
		/// Remove the first column from a Table.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Remove the first column from this table.
		///     table.RemoveColumn(0);
		///
		///     // Save the document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void RemoveColumn(int index)
		{
			if (index < 0 || index > ColumnCount - 1)
				throw new IndexOutOfRangeException();

			foreach (Row r in Rows)
				if (r.Cells.Count < ColumnCount)
				{
					var positionIndex = 0;
					var actualPosition = 0;
					var gridAfterVal = 0;
					// checks to see if there is a deleted cell                    
					gridAfterVal = r.gridAfter;

					// goes through iteration of cells to find the one the that contains the index number
					foreach (Cell rowCell in r.Cells)
					{
						// checks if the cell has a gridspan
						var gridSpanVal = 0;

						if (rowCell.GridSpan != 0)
						{
							gridSpanVal = rowCell.GridSpan - 1;
						}

						// checks to see if the index is within its lowest and highest cell value
						if ((index - gridAfterVal) >= actualPosition
							&& (index - gridAfterVal) <= (actualPosition + gridSpanVal))
						{
							r.Cells[positionIndex].Xml.Remove();
							break;
						}
						positionIndex += 1;
						actualPosition += gridSpanVal + 1;
					}
				}
				else
				{
					r.Cells[index].Xml.Remove();
				}

			_cachedColCount = -1;
		}

		/// <summary>
		/// Insert a row into this table.
		/// </summary>
		/// <example>
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first table in this document.
		///     Table table = document.Tables[0];
		///        
		///     // Insert a new row at index 1 in this table.
		///     Row row = table.InsertRow(1);
		///
		///     // Loop through each cell in this new row.
		///     foreach (Cell c in row.Cells)
		///     {
		///         // Set the text of each new cell to "Hello".
		///         c.Paragraphs[0].InsertText("Hello", false);
		///     }
		///
		///     // Save the document to a new file.
		///     document.SaveAs(@"C:\Example\Test2.docx");
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		/// <param name="index">Index to insert row at.</param>
		/// <returns>A new Row</returns>
		public Row InsertRow(int index)
		{
			if (index < 0 || index > RowCount)
				throw new IndexOutOfRangeException();

			List<XElement> content = new List<XElement>();
			for (int i = 0; i < ColumnCount; i++)
			{
				var w = 2310d;
				if (ColumnWidthsValue != null && ColumnWidthsValue.Length > i)
					w = ColumnWidthsValue[i] * 15;
				XElement cell = HelperFunctions.CreateTableCell(w);
				content.Add(cell);
			}

			return InsertRow(content, index);
		}

		/// <summary>
		/// Insert a copy of a row into this table.
		/// </summary>
		/// <param name="row">Row to copy and insert.</param>
		/// <param name="index">Index to insert row at.</param>
		/// <returns>A new Row</returns>
		public Row InsertRow(Row row, int index)
		{
			if (row == null)
				throw new ArgumentNullException(nameof(row));

			if (index < 0 || index > RowCount)
				throw new IndexOutOfRangeException();

			List<XElement> content = row.Xml.Elements(XName.Get("tc", DocX.w.NamespaceName)).Select(element => HelperFunctions.CloneElement(element)).ToList();
			return InsertRow(content, index);
		}

		private Row InsertRow(List<XElement> content, Int32 index)
		{
			Row newRow = new Row(this, Document, new XElement(XName.Get("tr", DocX.w.NamespaceName), content));

			XElement rowXml;
			if (index == Rows.Count)
			{
				rowXml = Rows.Last().Xml;
				rowXml.AddAfterSelf(newRow.Xml);
			}

			else
			{
				rowXml = Rows[index].Xml;
				rowXml.AddBeforeSelf(newRow.Xml);
			}

			return newRow;
		}

		/// <summary>
		/// Insert a column into a table.
		/// </summary>
		/// <param name="index">The index to insert the column at.</param>
		/// <param name="direction">The side in which you wish to place the colum(True right, false left)</param>
		/// <example>
		/// Insert a column to the left of a table.
		/// <code>
		/// // Load a document.
		/// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		/// {
		///     // Get the first Table in this document.
		///     Table table = document.Tables[0];
		///
		///     // Insert a new column to this left of this table.
		///     table.InsertColumn(0, false);
		///
		///     // Set the new columns text to "Row no."
		///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
		///
		///     // Loop through each row in the table.
		///     for (int i = 1; i &lt; table.Rows.Count; i++)
		///     {
		///         // The current row.
		///         Row row = table.Rows[i];
		///
		///         // The cell in this row that belongs to the new column.
		///         Cell cell = row.Cells[table.ColumnCount - 1];
		///
		///         // The first Paragraph that this cell houses.
		///         Paragraph p = cell.Paragraphs[0];
		///
		///         // Insert this rows index.
		///         p.InsertText(i.ToString(), false);
		///     }
		///
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public void InsertColumn(int index, bool direction)
		{
			var columnCount = ColumnCount;
			if (RowCount > 0)
			{
				if (index > 0 && index <= columnCount)
				{
					_cachedColCount = -1;
					foreach (Row r in Rows)
					{
						// create cell
						XElement cell = HelperFunctions.CreateTableCell();

						// insert cell 
						// checks if it is in bounds of index
						if (r.Cells.Count < columnCount)
						{
							if (index >= columnCount)
							{
								AddCellToRow(r, cell, r.Cells.Count, direction);
							}
							else
							{
								bool directionTest = true;
								var positionIndex = 1;
								var actualPosition = 1;
								var gridAfterVal = 0;
								// checks to see if there is a deleted cell

								gridAfterVal = r.gridAfter;

								// goes through iteration of cells to find the one the that contains the index number
								foreach (Cell rowCell in r.Cells)
								{
									// checks if the cell has a gridspan
									var gridSpanVal = 0;

									if (rowCell.GridSpan != 0)
									{
										gridSpanVal = rowCell.GridSpan - 1;
									}

									// checks to see if the index is within its lowest and highest cell value
									if ((index - gridAfterVal) >= actualPosition
										&& (index - gridAfterVal) <= (actualPosition + gridSpanVal))
									{
										if (index == (actualPosition + gridSpanVal) && direction == true)
										{
											directionTest = true;
										}
										else
										{
											directionTest = false;
										}
										AddCellToRow(r, cell, positionIndex, directionTest);
										break;
									}
									positionIndex += 1;
									actualPosition += gridSpanVal + 1;
								}
							}
						}
						else if (r.Cells.Count == index)
						{
							AddCellToRow(r, cell, index, direction);
						}
						else
							AddCellToRow(r, cell, index, direction);
					}
				}
				else
				{
					throw new IndexOutOfRangeException("Out of index bounds, column count is " + columnCount + " you input " + index);
				}
			}
		}

		/// <summary>
		/// Adds a cell to the right or left of a cell
		/// </summary>
		/// <param name="row">is the row you are adding</param>
		/// <param name="cell">is the cell you are adding</param>
		/// <param name="index">the cell index position you are refferencing from</param>
		/// <param name="direction">which side of the cell you wish to add cell</param>

		private void AddCellToRow(Row row, XElement cell, int index, bool direction)
		{
			index -= 1;
			if (direction)
			{
				row.Cells[index].Xml.AddAfterSelf(cell);
			}
			else
			{
				row.Cells[index].Xml.AddBeforeSelf(cell);
			}
		}
		/// <summary>
		/// Deletes a cell in a row
		/// </summary>
		/// <param name="rowIndex">index of the row you want to remove the cell</param>
		/// <param name="celIndex">index of the cell you want to remove</param>
		public void DeleteAndShiftCellsLeft(int rowIndex, int celIndex)
		{

			XAttribute gridAfterVal = new XAttribute(XName.Get("val", DocX.w.NamespaceName), 0);
			var trPr = Rows[rowIndex].Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			if (trPr != null)
			{
				var gridAfter = trPr.Element(XName.Get("gridAfter", DocX.w.NamespaceName));
				if (gridAfter != null)
				{
					var val = gridAfter.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (val != null)
					{
						val.Value = (int.Parse(val.Value) + 1).ToString();
					}
					else
					{
						val.Value = "1";
					}
				}
				else
				{
					var gridAfterElement = new XElement("gridAfter");
					var gridAfterValAttribute = new XAttribute("val", 1);
					gridAfter.SetAttributeValue("val", 1);
				}
			}
			else
			{
				XElement trPrXElement = new XElement(XName.Get("trPr", DocX.w.NamespaceName));
				XElement gridAfterElement = new XElement(XName.Get("gridAfter", DocX.w.NamespaceName));
				XAttribute gridAfterValAttribute = new XAttribute(XName.Get("val", DocX.w.NamespaceName), 1);
				gridAfterElement.Add(gridAfterValAttribute);
				trPrXElement.Add(gridAfterElement);
				Rows[rowIndex].Xml.AddFirst(trPrXElement);
			}
			var columnCount = this.ColumnCount;
			if (celIndex <= this.ColumnCount && this.Rows[rowIndex].ColumnCount <= this.ColumnCount)
			{
				Rows[rowIndex].Cells[celIndex].Xml.Remove();
			}
		}

		/// <summary>
		/// Insert a page break before a Table.
		/// </summary>
		/// <example>
		/// Insert a Table and a Paragraph into a document with a page break between them.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {              
		///     // Insert a new Paragraph.
		///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
		///
		///     // Insert a new Table.
		///     Table t1 = document.InsertTable(2, 2);
		///     t1.Design = TableDesign.LightShadingAccent1;
		///     
		///     // Insert a page break before this Table.
		///     t1.InsertPageBreakBeforeSelf();
		///     
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override void InsertPageBreakBeforeSelf()
		{
			base.InsertPageBreakBeforeSelf();
		}


		/// <summary>
		/// Insert a page break after a Table.
		/// </summary>
		/// <example>
		/// Insert a Table and a Paragraph into a document with a page break between them.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a new Table.
		///     Table t1 = document.InsertTable(2, 2);
		///     t1.Design = TableDesign.LightShadingAccent1;
		///        
		///     // Insert a page break after this Table.
		///     t1.InsertPageBreakAfterSelf();
		///        
		///     // Insert a new Paragraph.
		///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
		///
		///     // Save this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override void InsertPageBreakAfterSelf()
		{
			base.InsertPageBreakAfterSelf();
		}

		/// <summary>
		/// Insert a new Table before this Table, this Table can be from this document or another document.
		/// </summary>
		/// <param name="t">The Table t to be inserted</param>
		/// <returns>A new Table inserted before this Table.</returns>
		/// <example>
		/// Insert a new Table before this Table.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Table in document b.
		///     Table t2 = documentB.Tables[0];
		///
		///     // Insert the Table from document a before this Table.
		///     Table newTable = t2.InsertTableBeforeSelf(t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableBeforeSelf(Table t)
		{
			return base.InsertTableBeforeSelf(t);
		}

		/// <summary>
		/// Insert a new Table into this document before this Table.
		/// </summary>
		/// <param name="rowCount">The number of rows this Table should have.</param>
		/// <param name="columnCount">The number of columns this Table should have.</param>
		/// <returns>A new Table inserted before this Table.</returns>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     //Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///     t.Design = TableDesign.LightShadingAccent1;
		///     t.Alignment = Alignment.center;
		///     
		///     // Insert a new Table before this Table.
		///     Table newTable = t.InsertTableBeforeSelf(2, 2);
		///     newTable.Design = TableDesign.LightShadingAccent2;
		///     newTable.Alignment = Alignment.center;
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableBeforeSelf(int rowCount, int columnCount)
		{
			return base.InsertTableBeforeSelf(rowCount, columnCount);
		}

		/// <summary>
		/// Insert a new Table after this Table, this Table can be from this document or another document.
		/// </summary>
		/// <param name="t">The Table t to be inserted</param>
		/// <returns>A new Table inserted after this Table.</returns>
		/// <example>
		/// Insert a new Table after this Table.
		/// <code>
		/// // Place holder for a Table.
		/// Table t;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first Table from this document.
		///     t = documentA.Tables[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Table in document b.
		///     Table t2 = documentB.Tables[0];
		///
		///     // Insert the Table from document a after this Table.
		///     Table newTable = t2.InsertTableAfterSelf(t);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableAfterSelf(Table t)
		{
			return base.InsertTableAfterSelf(t);
		}

		/// <summary>
		/// Insert a new Table into this document after this Table.
		/// </summary>
		/// <param name="rowCount">The number of rows this Table should have.</param>
		/// <param name="columnCount">The number of columns this Table should have.</param>
		/// <returns>A new Table inserted before this Table.</returns>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     //Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///     t.Design = TableDesign.LightShadingAccent1;
		///     t.Alignment = Alignment.center;
		///     
		///     // Insert a new Table after this Table.
		///     Table newTable = t.InsertTableAfterSelf(2, 2);
		///     newTable.Design = TableDesign.LightShadingAccent2;
		///     newTable.Alignment = Alignment.center;
		///
		///     // Save all changes made to this document.
		///     document.Save();
		/// }// Release this document from memory.
		/// </code>
		/// </example>
		public override Table InsertTableAfterSelf(int rowCount, int columnCount)
		{
			return base.InsertTableAfterSelf(rowCount, columnCount);
		}

		/// <summary>
		/// Insert a Paragraph before this Table, this Paragraph may have come from the same or another document.
		/// </summary>
		/// <param name="p">The Paragraph to insert.</param>
		/// <returns>The Paragraph now associated with this document.</returns>
		/// <example>
		/// Take a Paragraph from document a, and insert it into document b before this Table.
		/// <code>
		/// // Place holder for a Paragraph.
		/// Paragraph p;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first paragraph from this document.
		///     p = documentA.Paragraphs[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Table in document b.
		///     Table t = documentB.Tables[0];
		///
		///     // Insert the Paragraph from document a before this Table.
		///     Paragraph newParagraph = t.InsertParagraphBeforeSelf(p);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code> 
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
		{
			return base.InsertParagraphBeforeSelf(p);
		}

		/// <summary>
		/// Insert a new Paragraph before this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <returns>A new Paragraph inserted before this Table.</returns>
		/// <example>
		/// Insert a new Paragraph before the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.");
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text)
		{
			return base.InsertParagraphBeforeSelf(text);
		}

		/// <summary>
		/// Insert a new Paragraph before this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <returns>A new Paragraph inserted before this Table.</returns>
		/// <example>
		/// Insert a new paragraph before the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
		{
			return base.InsertParagraphBeforeSelf(text, trackChanges);
		}

		/// <summary>
		/// Insert a new Paragraph before this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <param name="formatting">The formatting to apply to this insertion.</param>
		/// <returns>A new Paragraph inserted before this Table.</returns>
		/// <example>
		/// Insert a new paragraph before the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     Formatting boldFormatting = new Formatting();
		///     boldFormatting.Bold = true;
		///
		///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false, boldFormatting);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
		{
			return base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
		}

		/// <summary>
		/// Insert a Paragraph after this Table, this Paragraph may have come from the same or another document.
		/// </summary>
		/// <param name="p">The Paragraph to insert.</param>
		/// <returns>The Paragraph now associated with this document.</returns>
		/// <example>
		/// Take a Paragraph from document a, and insert it into document b after this Table.
		/// <code>
		/// // Place holder for a Paragraph.
		/// Paragraph p;
		///
		/// // Load document a.
		/// using (DocX documentA = DocX.Load(@"a.docx"))
		/// {
		///     // Get the first paragraph from this document.
		///     p = documentA.Paragraphs[0];
		/// }
		///
		/// // Load document b.
		/// using (DocX documentB = DocX.Load(@"b.docx"))
		/// {
		///     // Get the first Table in document b.
		///     Table t = documentB.Tables[0];
		///
		///     // Insert the Paragraph from document a after this Table.
		///     Paragraph newParagraph = t.InsertParagraphAfterSelf(p);
		///
		///     // Save all changes made to document b.
		///     documentB.Save();
		/// }// Release this document from memory.
		/// </code> 
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(Paragraph p)
		{
			return base.InsertParagraphAfterSelf(p);
		}

		/// <summary>
		/// Insert a new Paragraph after this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <param name="formatting">The formatting to apply to this insertion.</param>
		/// <returns>A new Paragraph inserted after this Table.</returns>
		/// <example>
		/// Insert a new paragraph after the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     Formatting boldFormatting = new Formatting();
		///     boldFormatting.Bold = true;
		///
		///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false, boldFormatting);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
		{
			return base.InsertParagraphAfterSelf(text, trackChanges, formatting);
		}

		/// <summary>
		/// Insert a new Paragraph after this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <param name="trackChanges">Should this insertion be tracked as a change?</param>
		/// <returns>A new Paragraph inserted after this Table.</returns>
		/// <example>
		/// Insert a new paragraph after the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false);
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
		{
			return base.InsertParagraphAfterSelf(text, trackChanges);
		}

		/// <summary>
		/// Insert a new Paragraph after this Table.
		/// </summary>
		/// <param name="text">The initial text for this new Paragraph.</param>
		/// <returns>A new Paragraph inserted after this Table.</returns>
		/// <example>
		/// Insert a new Paragraph after the first Table in this document.
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create(@"Test.docx"))
		/// {
		///     // Insert a Table into this document.
		///     Table t = document.InsertTable(2, 2);
		///
		///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.");
		///
		///     // Save all changes made to this new document.
		///     document.Save();
		///    }// Release this new document form memory.
		/// </code>
		/// </example>
		public override Paragraph InsertParagraphAfterSelf(string text)
		{
			return base.InsertParagraphAfterSelf(text);
		}

		/// <summary>
		/// Set a table border
		/// Added by lckuiper @ 20101117
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert a table into this document.
		///    Table t = document.InsertTable(3, 3);
		///
		///    // Create a large blue border.
		///    Border b = new Border(BorderStyle.Tcbs_single, BorderSize.seven, 0, Color.Blue);
		///
		///    // Set the tables Top, Bottom, Left and Right Borders to b.
		///    t.SetBorder(TableBorderType.Top, b);
		///    t.SetBorder(TableBorderType.Bottom, b);
		///    t.SetBorder(TableBorderType.Left, b);
		///    t.SetBorder(TableBorderType.Right, b);
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		/// <param name="borderType">The table border to set</param>
		/// <param name="border">Border object to set the table border</param>
		public void SetBorder(TableBorderType borderType, Border border)
		{
			/*
			 * Get the tblPr (table properties) element for this Table,
			 * null will be return if no such element exists.
			 */
			XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			if (tblPr == null)
			{
				Xml.SetElementValue(XName.Get("tblPr", DocX.w.NamespaceName), string.Empty);
				tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			}

			/*
			 * Get the tblBorders (table borders) element for this Table,
			 * null will be return if no such element exists.
			 */
			XElement tblBorders = tblPr.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			if (tblBorders == null)
			{
				tblPr.SetElementValue(XName.Get("tblBorders", DocX.w.NamespaceName), string.Empty);
				tblBorders = tblPr.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			}

			/*
			 * Get the 'borderType' (table border) element for this Table,
			 * null will be return if no such element exists.
			 */
			string tbordertype;
			tbordertype = borderType.ToString();
			// only lower the first char of string (because of insideH and insideV)
			tbordertype = tbordertype.Substring(0, 1).ToLower() + tbordertype.Substring(1);

			XElement tblBorderType = tblBorders.Element(XName.Get(borderType.ToString(), DocX.w.NamespaceName));
			if (tblBorderType == null)
			{
				tblBorders.SetElementValue(XName.Get(tbordertype, DocX.w.NamespaceName), string.Empty);
				tblBorderType = tblBorders.Element(XName.Get(tbordertype, DocX.w.NamespaceName));
			}

			// get string value of border style
			string borderstyle = border.Tcbs.ToString().Substring(5);
			borderstyle = borderstyle.Substring(0, 1).ToLower() + borderstyle.Substring(1);

			// The val attribute is used for the border style
			tblBorderType.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), borderstyle);

			if (border.Tcbs != BorderStyle.Tcbs_nil)
			{
				int size;
				switch (border.Size)
				{
					case BorderSize.one: size = 2; break;
					case BorderSize.two: size = 4; break;
					case BorderSize.three: size = 6; break;
					case BorderSize.four: size = 8; break;
					case BorderSize.five: size = 12; break;
					case BorderSize.six: size = 18; break;
					case BorderSize.seven: size = 24; break;
					case BorderSize.eight: size = 36; break;
					case BorderSize.nine: size = 48; break;
					default: size = 2; break;
				}

				// The sz attribute is used for the border size
				tblBorderType.SetAttributeValue(XName.Get("sz", DocX.w.NamespaceName), (size).ToString());

				// The space attribute is used for the cell spacing (probably '0')
				tblBorderType.SetAttributeValue(XName.Get("space", DocX.w.NamespaceName), (border.Space).ToString());

				// The color attribute is used for the border color
				tblBorderType.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), border.Color.ToHex());
			}
		}

		/// <summary>
		/// Get a table border
		/// Added by lckuiper @ 20101117
		/// </summary>
		/// <param name="borderType">The table border to get</param>
		public Border GetBorder(TableBorderType borderType)
		{
			// instance with default border values
			Border b = new Border();

			/*
             * Get the tblPr (table properties) element for this Table,
             * null will be return if no such element exists.
             */
			XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			if (tblPr == null)
			{
				// uses default border style
			}

			/*
             * Get the tblBorders (table borders) element for this Table,
             * null will be return if no such element exists.
             */
			XElement tblBorders = tblPr.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			if (tblBorders == null)
			{
				// uses default border style
			}

			/*
             * Get the 'borderType' (table border) element for this Table,
             * null will be return if no such element exists.
             */
			string tbordertype;
			tbordertype = borderType.ToString();
			// only lower the first char of string (because of insideH and insideV)
			tbordertype = tbordertype.Substring(0, 1).ToLower() + tbordertype.Substring(1);

			XElement tblBorderType = tblBorders.Element(XName.Get(tbordertype, DocX.w.NamespaceName));
			if (tblBorderType == null)
			{
				// uses default border style
			}

			// The val attribute is used for the border style
			XAttribute val = tblBorderType.Attribute(XName.Get("val", DocX.w.NamespaceName));
			// If val is null, this table contains no border information.
			if (val == null)
			{
				// uses default border style
			}
			else
			{
				try
				{
					string bordertype = "Tcbs_" + val.Value;
					b.Tcbs = (BorderStyle)Enum.Parse(typeof(BorderStyle), bordertype);
				}
				catch
				{
					val.Remove();
					// uses default border style
				}
			}

			// The sz attribute is used for the border size
			XAttribute sz = tblBorderType.Attribute(XName.Get("sz", DocX.w.NamespaceName));
			// If sz is null, this border contains no size information.
			if (sz == null)
			{
				// uses default border style
			}
			else
			{
				// If sz is not an int, something is wrong with this attributes value, so remove it
				int numerical_size;
				if (!int.TryParse(sz.Value, out numerical_size))
					sz.Remove();
				else
				{
					switch (numerical_size)
					{
						case 2: b.Size = BorderSize.one; break;
						case 4: b.Size = BorderSize.two; break;
						case 6: b.Size = BorderSize.three; break;
						case 8: b.Size = BorderSize.four; break;
						case 12: b.Size = BorderSize.five; break;
						case 18: b.Size = BorderSize.six; break;
						case 24: b.Size = BorderSize.seven; break;
						case 36: b.Size = BorderSize.eight; break;
						case 48: b.Size = BorderSize.nine; break;
						default: b.Size = BorderSize.one; break;
					}
				}
			}

			// The space attribute is used for the border spacing (probably '0')
			XAttribute space = tblBorderType.Attribute(XName.Get("space", DocX.w.NamespaceName));
			// If space is null, this border contains no space information.
			if (space == null)
			{
				// uses default border style
			}
			else
			{
				// If space is not an int, something is wrong with this attributes value, so remove it
				int borderspace;
				if (!int.TryParse(space.Value, out borderspace))
				{
					space.Remove();
					// uses default border style
				}
				else
				{
					b.Space = borderspace;
				}
			}

			// The color attribute is used for the border color
			XAttribute color = tblBorderType.Attribute(XName.Get("color", DocX.w.NamespaceName));
			if (color == null)
			{
				// uses default border style
			}
			else
			{
				// If color is not a Color, something is wrong with this attributes value, so remove it
				try
				{
					b.Color = ColorTranslator.FromHtml(string.Format("#{0}", color.Value));
				}
				catch
				{
					color.Remove();
					// uses default border style
				}
			}
			return b;
		}

	}

	/// <summary>Represents a single row in a Table</summary>
	public class Row : Container
	{
		/// <summary>
		/// Calculates columns count in the row, taking spanned cells into account
		/// </summary>
		public Int32 ColumnCount
		{
			get
			{
				int gridSpanSum = 0;

				gridSpanSum += gridAfter;

				// Foreach each Cell between startIndex and endIndex inclusive.
				foreach (Cell c in Cells)
				{
					if (c.GridSpan != 0)
					{
						gridSpanSum += c.GridSpan - 1;
					}
				}

				// return cells count + count of spanned cells
				return Cells.Count + gridSpanSum;
			}
		}

		/// <summary>
		/// Returns the GridAfter of a row ie. The amount of cells that are deleted
		/// </summary>
		public int gridAfter
		{
			get
			{
				var gridAfterValue = 0;
				var trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				if (trPr != null)
				{
					var gridAfter = trPr.Element(XName.Get("gridAfter", DocX.w.NamespaceName));
					if (gridAfter != null)
					{
						var val = gridAfter.Attribute(XName.Get("val", DocX.w.NamespaceName));
						if (val != null)
						{
							gridAfterValue += int.Parse(val.Value);
						}
					}
				}
				return gridAfterValue;
			}
		}

		/// <summary>
		/// A list of Cells in this Row.
		/// </summary>
		public List<Cell> Cells
		{
			get
			{
				List<Cell> cells =
				(
					from c in Xml.Elements(XName.Get("tc", DocX.w.NamespaceName))
					select new Cell(this, Document, c)
				).ToList();

				return cells;
			}
		}

		/// <summary></summary>
		public void Remove()
		{
			XElement table = Xml.Parent;

			Xml.Remove();
			if (table.Elements(XName.Get("tr", DocX.w.NamespaceName)).Count() == 0)
				table.Remove();
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				List<Paragraph> paragraphs =
				(
					from p in Xml.Descendants(DocX.w + "p")
					select new Paragraph(Document, p, 0)
				).ToList();

				foreach (Paragraph p in paragraphs)
					p.PackagePart = table.mainPart;

				return paragraphs.AsReadOnly();
			}
		}

		internal Table table;
		internal Row(Table table, DocX document, XElement xml)
			: base(document, xml)
		{
			this.table = table;
			this.mainPart = table.mainPart;
		}

		/// <summary>
		/// The property name to set when specifiying an exact height
		/// </summary>
		/// <created>Nick Kusters</created>
		const string _hRule_Exact = "exact";
		/// <summary>
		/// The property name to set when specifying a minimum height
		/// </summary>
		/// <created>Nick Kusters</created>
		const string _hRule_AtLeast = "atLeast";
		/// <summary>
		/// Height in pixels. // Added by Joel, refactored by Cathal.
		/// </summary>
		public double Height
		{
			get
			{
				/*
                * Get the trPr (table row properties) element for this Row,
                * null will be return if no such element exists.
                */
				XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));

				// If trPr is null, this row contains no height information.
				if (trPr == null)
					return double.NaN;

				/*
                 * Get the trHeight element for this Row,
                 * null will be return if no such element exists.
                 */
				XElement trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));

				// If trHeight is null, this row contains no height information.
				if (trHeight == null)
					return double.NaN;

				// Get the val attribute for this trHeight element.
				XAttribute val = trHeight.Attribute(XName.Get("val", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (val == null)
					return double.NaN;

				// If val is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double heightInWordUnits;
				if (!double.TryParse(val.Value, out heightInWordUnits))
				{
					val.Remove();
					return double.NaN;
				}

				// 15 "word units" in one pixel
				return (heightInWordUnits / 15);
			}
			set
			{
				SetHeight(value, true);
			}
		}
		/// <summary>
		/// Helper method to set either the exact height or the min-height
		/// </summary>
		/// <param name="height">The height value to set (in pixels)</param>
		/// <param name="exact">
		/// If true, the height will be forced. 
		/// If false, it will be treated as a minimum height, auto growing past it if need be.
		/// </param>
		/// <created>Nick Kusters</created>
		void SetHeight(double height, bool exact)
		{
			/*
             * Get the trPr (table row properties) element for this Row,
             * null will be return if no such element exists.
             */
			XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			if (trPr == null)
			{
				Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
				trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			}

			/*
             * Get the trHeight element for this Row,
             * null will be return if no such element exists.
             */
			XElement trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));
			if (trHeight == null)
			{
				trPr.SetElementValue(XName.Get("trHeight", DocX.w.NamespaceName), string.Empty);
				trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));
			}

			// The hRule attribute needs to be set to exact.
			trHeight.SetAttributeValue(XName.Get("hRule", DocX.w.NamespaceName), exact ? _hRule_Exact : _hRule_AtLeast);

			// 15 "word units" is equal to one pixel. 
			trHeight.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (height * 15).ToString());
		}
		/// <summary>
		/// Min-Height in pixels. // Added by Nick Kusters.
		/// </summary>
		/// <remarks>
		/// Value will be treated as a minimum height, auto growing past it if need be.
		/// </remarks>
		/// <created>Nick Kusters</created>
		public double MinHeight
		{
			get
			{
				// Just return the value from the normal height property since it doesn't care if you've set an exact or minimum height.
				return Height;
			}
			set
			{
				SetHeight(value, false);
			}
		}


		/// <summary>
		/// Set to true to make this row the table header row that will be repeated on each page
		/// </summary>
		public bool TableHeader
		{
			get
			{
				XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				XElement tblHeader = trPr.Element(XName.Get("tblHeader", DocX.w.NamespaceName));
				if (tblHeader == null)
				{
					return false;
				}
				else
				{
					return true;
				}
			}
			set
			{
				XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				if (trPr == null)
				{
					Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
					trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				}
				XElement tblHeader = trPr.Element(XName.Get("tblHeader", DocX.w.NamespaceName));
				if (tblHeader == null && value)
				{
					trPr.SetElementValue(XName.Get("tblHeader", DocX.w.NamespaceName), string.Empty);
				}
				if (tblHeader != null && !value)
				{
					tblHeader.Remove();
				}
			}
		}


		/// <summary>
		/// Allow row to break across pages. 
		/// The default value is true: Word will break the contents of the row across pages. 
		/// If set to false, the contents of the row will not be split across pages, the entire row will be moved to the next page instead.
		/// </summary>
		public bool BreakAcrossPages
		{
			get
			{
				XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));

				if (trPr == null)
					return true;

				XElement trCantSplit = trPr.Element(XName.Get("cantSplit", DocX.w.NamespaceName));

				if (trCantSplit == null)
					return true;

				return false;
			}

			set
			{
				if (value == false)
				{
					XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
					if (trPr == null)
					{
						Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
						trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
					}

					XElement trCantSplit = trPr.Element(XName.Get("cantSplit", DocX.w.NamespaceName));
					if (trCantSplit == null)
						trPr.SetElementValue(XName.Get("cantSplit", DocX.w.NamespaceName), string.Empty);
				}

				if (value == true)
				{
					XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
					if (trPr != null)
					{
						XElement trCantSplit = trPr.Element(XName.Get("cantSplit", DocX.w.NamespaceName));
						if (trCantSplit != null)
							trCantSplit.Remove();
					}
				}
			}
		}

		/// <summary>
		/// Merge cells starting with startIndex and ending with endIndex.
		/// </summary>
		public void MergeCells(int startIndex, int endIndex)
		{
			// Check for valid start and end indexes.
			if (startIndex < 0 || endIndex <= startIndex || endIndex > Cells.Count + 1)
				throw new IndexOutOfRangeException();

			// The sum of all merged gridSpans.
			int gridSpanSum = 0;

			// Foreach each Cell between startIndex and endIndex inclusive.
			foreach (Cell c in Cells.Where((z, i) => i > startIndex && i <= endIndex))
			{
				XElement tcPr = c.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr != null)
				{
					XElement gridSpan = tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
					if (gridSpan != null)
					{
						XAttribute val = gridSpan.Attribute(XName.Get("val", DocX.w.NamespaceName));

						int value = 0;
						if (val != null)
							if (int.TryParse(val.Value, out value))
								gridSpanSum += value - 1;
					}
				}

				// Add this cells Pragraph to the merge start Cell.
				Cells[startIndex].Xml.Add(c.Xml.Elements(XName.Get("p", DocX.w.NamespaceName)));

				// Remove this Cell.
				c.Xml.Remove();
			}

			/* 
             * Get the tcPr (table cell properties) element for the first cell in this merge,
             * null will be returned if no such element exists.
             */
			XElement start_tcPr = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (start_tcPr == null)
			{
				Cells[startIndex].Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				start_tcPr = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}

			/* 
             * Get the gridSpan element of this row,
             * null will be returned if no such element exists.
             */
			XElement start_gridSpan = start_tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
			if (start_gridSpan == null)
			{
				start_tcPr.SetElementValue(XName.Get("gridSpan", DocX.w.NamespaceName), string.Empty);
				start_gridSpan = start_tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
			}

			/* 
             * Get the val attribute of this row,
             * null will be returned if no such element exists.
             */
			XAttribute start_val = start_gridSpan.Attribute(XName.Get("val", DocX.w.NamespaceName));

			int start_value = 0;
			if (start_val != null)
				if (int.TryParse(start_val.Value, out start_value))
					gridSpanSum += start_value - 1;

			// Set the val attribute to the number of merged cells.
			start_gridSpan.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (gridSpanSum + (endIndex - startIndex + 1)).ToString());
		}
	}

	/// <summary></summary>
	public class Cell : Container
	{

		internal Row row;

		internal Cell(Row row, DocX document, XElement xml) : base(document, xml)
		{
			this.row = row;
			this.mainPart = row.mainPart;
		}

		/// <summary></summary>
		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> paragraphs = base.Paragraphs;
				foreach (Paragraph p in paragraphs) p.PackagePart = row.table.mainPart;
				return paragraphs;
			}
		}

		/// <summary>Returns the GridSpan of a specific Cell ie. How many cells are merged.</summary>
		public int GridSpan
		{
			get
			{
				var gridSpanVal = 0;
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr != null)
				{
					XElement gridSpan = tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
					if (gridSpan != null)
					{
						XAttribute val = gridSpan.Attribute(XName.Get("val", DocX.w.NamespaceName));

						int value = 0;
						if (val != null)
							if (int.TryParse(val.Value, out value))
								gridSpanVal = value;
					}
				}
				return gridSpanVal;
			}
		}

		/// <summary>
		/// Gets or Sets this Cells vertical alignment.
		/// </summary>
		/// <!--Patch 7398 added by lckuiper on Nov 16th 2010 @ 2:23 PM-->
		/// <example>
		/// Creates a table with 3 cells and sets the vertical alignment of each to 1 of the 3 available options.
		/// <code>
		/// // Create a new document.
		///using(DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert a Table into this document.
		///    Table t = document.InsertTable(3, 1);
		///
		///    // Set the design of the Table such that we can easily identify cell boundaries.
		///    t.Design = TableDesign.TableGrid;
		///
		///    // Set the height of the row bigger than default.
		///    // We need to be able to see the difference in vertical cell alignment options.
		///    t.Rows[0].Height = 100;
		///
		///    // Set the vertical alignment of cell0 to top.
		///    Cell c0 = t.Rows[0].Cells[0];
		///    c0.InsertParagraph("VerticalAlignment.Top");
		///    c0.VerticalAlignment = VerticalAlignment.Top;
		///
		///    // Set the vertical alignment of cell1 to center.
		///    Cell c1 = t.Rows[0].Cells[1];
		///    c1.InsertParagraph("VerticalAlignment.Center");
		///    c1.VerticalAlignment = VerticalAlignment.Center;
		///
		///    // Set the vertical alignment of cell2 to bottom.
		///    Cell c2 = t.Rows[0].Cells[2];
		///    c2.InsertParagraph("VerticalAlignment.Bottom");
		///    c2.VerticalAlignment = VerticalAlignment.Bottom;
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public VerticalAlignment VerticalAlignment
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return VerticalAlignment.Center;

				/*
                 * Get the vAlign (table cell vertical alignment) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement vAlign = tcPr.Element(XName.Get("vAlign", DocX.w.NamespaceName));

				// If vAlign is null, this cell contains no vertical alignment information.
				if (vAlign == null)
					return VerticalAlignment.Center;

				// Get the val attribute of the vAlign element.
				XAttribute val = vAlign.Attribute(XName.Get("val", DocX.w.NamespaceName));

				// If val is null, this cell contains no vAlign information.
				if (val == null)
					return VerticalAlignment.Center;

				// If val is not a VerticalAlign enum, something is wrong with this attributes value, so remove it and return VerticalAlignment.Center;
				try
				{
					return (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), val.Value, true);
				}

				catch
				{
					val.Remove();
					return VerticalAlignment.Center;
				}
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the vAlign (table cell vertical alignment) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement vAlign = tcPr.Element(XName.Get("vAlign", DocX.w.NamespaceName));
				if (vAlign == null)
				{
					tcPr.SetElementValue(XName.Get("vAlign", DocX.w.NamespaceName), string.Empty);
					vAlign = tcPr.Element(XName.Get("vAlign", DocX.w.NamespaceName));
				}

				// Set the VerticalAlignment in 'val'
				vAlign.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value.ToString().ToLower());
			}
		}

		/// <summary></summary>
		public Color Shading
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no Color information.
				if (tcPr == null)
					return Color.White;

				/*
                 * Get the shd (table shade) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));

				// If shd is null, this cell contains no Color information.
				if (shd == null)
					return Color.White;

				// Get the w attribute of the tcW element.
				XAttribute fill = shd.Attribute(XName.Get("fill", DocX.w.NamespaceName));

				// If fill is null, this cell contains no Color information.
				if (fill == null)
					return Color.White;

				return ColorTranslator.FromHtml(string.Format("#{0}", fill.Value));
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the shd (table shade) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
				if (shd == null)
				{
					tcPr.SetElementValue(XName.Get("shd", DocX.w.NamespaceName), string.Empty);
					shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
				}

				// The val attribute needs to be set to clear
				shd.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "clear");

				// The color attribute needs to be set to auto
				shd.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), "auto");

				// The fill attribute needs to be set to the hex for this Color.
				shd.SetAttributeValue(XName.Get("fill", DocX.w.NamespaceName), value.ToHex());
			}
		}

		/// <summary>
		/// Width in pixels. // Added by Joel, refactored by Cathal
		/// </summary>
		public double Width
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return double.NaN;

				/*
                 * Get the tcW (table cell width) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));

				// If tcW is null, this cell contains no width information.
				if (tcW == null)
					return double.NaN;

				// Get the w attribute of the tcW element.
				XAttribute w = tcW.Attribute(XName.Get("w", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (w == null)
					return double.NaN;

				// If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double widthInWordUnits;
				if (!double.TryParse(w.Value, out widthInWordUnits))
				{
					w.Remove();
					return double.NaN;
				}

				// 15 "word units" is equal to one pixel.
				return (widthInWordUnits / 15);
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcW (table cell width) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));
				if (tcW == null)
				{
					tcPr.SetElementValue(XName.Get("tcW", DocX.w.NamespaceName), string.Empty);
					tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));
				}

				if (value == -1)
				{
					// remove cell width; due to set on table prop.
					tcW.Remove();
					return;

					//tcW.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "auto");
					//return;
				}

				// The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
				tcW.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

				// 15 "word units" is equal to one pixel. 
				tcW.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
			}
		}

		/// <summary>
		/// LeftMargin in pixels. // Added by lckuiper
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert table into this document.
		///    Table t = document.InsertTable(3, 3);
		///    t.Design = TableDesign.TableGrid;
		///
		///    // Get the center cell.
		///    Cell center = t.Rows[1].Cells[1];
		///
		///    // Insert some text so that we can see the effect of the Margins.
		///    center.Paragraphs[0].Append("Center Cell");
		///
		///    // Set the center cells Left, Margin to 10.
		///    center.MarginLeft = 25;
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public double MarginLeft
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return double.NaN;

				/*
                 * Get the tcMar
                 * 
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));

				// If tcMar is null, this cell contains no margin information.
				if (tcMar == null)
					return double.NaN;

				// Get the left (LeftMargin) element
				XElement tcMarLeft = tcMar.Element(XName.Get("left", DocX.w.NamespaceName));

				// If tcMarLeft is null, this cell contains no left margin information.
				if (tcMarLeft == null)
					return double.NaN;

				// Get the w attribute of the tcMarLeft element.
				XAttribute w = tcMarLeft.Attribute(XName.Get("w", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (w == null)
					return double.NaN;

				// If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double leftMarginInWordUnits;
				if (!double.TryParse(w.Value, out leftMarginInWordUnits))
				{
					w.Remove();
					return double.NaN;
				}

				// 15 "word units" is equal to one pixel.
				return (leftMarginInWordUnits / 15);
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcMar (table cell margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (tcMar == null)
				{
					tcPr.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}

				/*
                 * Get the left (table cell left margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMarLeft = tcMar.Element(XName.Get("left", DocX.w.NamespaceName));
				if (tcMarLeft == null)
				{
					tcMar.SetElementValue(XName.Get("left", DocX.w.NamespaceName), string.Empty);
					tcMarLeft = tcMar.Element(XName.Get("left", DocX.w.NamespaceName));
				}

				// The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
				tcMarLeft.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

				// 15 "word units" is equal to one pixel. 
				tcMarLeft.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
			}
		}

		/// <summary>
		/// RightMargin in pixels. // Added by lckuiper
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert table into this document.
		///    Table t = document.InsertTable(3, 3);
		///    t.Design = TableDesign.TableGrid;
		///
		///    // Get the center cell.
		///    Cell center = t.Rows[1].Cells[1];
		///
		///    // Insert some text so that we can see the effect of the Margins.
		///    center.Paragraphs[0].Append("Center Cell");
		///
		///    // Set the center cells Right, Margin to 10.
		///    center.MarginRight = 25;
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public double MarginRight
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return double.NaN;

				/*
                 * Get the tcMar
                 * 
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));

				// If tcMar is null, this cell contains no margin information.
				if (tcMar == null)
					return double.NaN;

				// Get the right (RightMargin) element
				XElement tcMarRight = tcMar.Element(XName.Get("right", DocX.w.NamespaceName));

				// If tcMarRight is null, this cell contains no right margin information.
				if (tcMarRight == null)
					return double.NaN;

				// Get the w attribute of the tcMarRight element.
				XAttribute w = tcMarRight.Attribute(XName.Get("w", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (w == null)
					return double.NaN;

				// If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double rightMarginInWordUnits;
				if (!double.TryParse(w.Value, out rightMarginInWordUnits))
				{
					w.Remove();
					return double.NaN;
				}

				// 15 "word units" is equal to one pixel.
				return (rightMarginInWordUnits / 15);
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcMar (table cell margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (tcMar == null)
				{
					tcPr.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}

				/*
                 * Get the right (table cell right margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMarRight = tcMar.Element(XName.Get("right", DocX.w.NamespaceName));
				if (tcMarRight == null)
				{
					tcMar.SetElementValue(XName.Get("right", DocX.w.NamespaceName), string.Empty);
					tcMarRight = tcMar.Element(XName.Get("right", DocX.w.NamespaceName));
				}

				// The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
				tcMarRight.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

				// 15 "word units" is equal to one pixel. 
				tcMarRight.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
			}
		}

		/// <summary>
		/// TopMargin in pixels. // Added by lckuiper
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert table into this document.
		///    Table t = document.InsertTable(3, 3);
		///    t.Design = TableDesign.TableGrid;
		///
		///    // Get the center cell.
		///    Cell center = t.Rows[1].Cells[1];
		///
		///    // Insert some text so that we can see the effect of the Margins.
		///    center.Paragraphs[0].Append("Center Cell");
		///
		///    // Set the center cells Top, Margin to 10.
		///    center.MarginTop = 25;
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public double MarginTop
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return double.NaN;

				/*
                 * Get the tcMar
                 * 
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));

				// If tcMar is null, this cell contains no margin information.
				if (tcMar == null)
					return double.NaN;

				// Get the top (TopMargin) element
				XElement tcMarTop = tcMar.Element(XName.Get("top", DocX.w.NamespaceName));

				// If tcMarTop is null, this cell contains no top margin information.
				if (tcMarTop == null)
					return double.NaN;

				// Get the w attribute of the tcMarTop element.
				XAttribute w = tcMarTop.Attribute(XName.Get("w", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (w == null)
					return double.NaN;

				// If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double topMarginInWordUnits;
				if (!double.TryParse(w.Value, out topMarginInWordUnits))
				{
					w.Remove();
					return double.NaN;
				}

				// 15 "word units" is equal to one pixel.
				return (topMarginInWordUnits / 15);
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcMar (table cell margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (tcMar == null)
				{
					tcPr.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}

				/*
                 * Get the top (table cell top margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMarTop = tcMar.Element(XName.Get("top", DocX.w.NamespaceName));
				if (tcMarTop == null)
				{
					tcMar.SetElementValue(XName.Get("top", DocX.w.NamespaceName), string.Empty);
					tcMarTop = tcMar.Element(XName.Get("top", DocX.w.NamespaceName));
				}

				// The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
				tcMarTop.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

				// 15 "word units" is equal to one pixel. 
				tcMarTop.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
			}
		}

		/// <summary>
		/// BottomMargin in pixels. // Added by lckuiper
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert table into this document.
		///    Table t = document.InsertTable(3, 3);
		///    t.Design = TableDesign.TableGrid;
		///
		///    // Get the center cell.
		///    Cell center = t.Rows[1].Cells[1];
		///
		///    // Insert some text so that we can see the effect of the Margins.
		///    center.Paragraphs[0].Append("Center Cell");
		///
		///    // Set the center cells Top, Margin to 10.
		///    center.MarginBottom = 25;
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		public double MarginBottom
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return double.NaN;

				/*
                 * Get the tcMar
                 * 
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));

				// If tcMar is null, this cell contains no margin information.
				if (tcMar == null)
					return double.NaN;

				// Get the bottom (BottomMargin) element
				XElement tcMarBottom = tcMar.Element(XName.Get("bottom", DocX.w.NamespaceName));

				// If tcMarBottom is null, this cell contains no bottom margin information.
				if (tcMarBottom == null)
					return double.NaN;

				// Get the w attribute of the tcMarBottom element.
				XAttribute w = tcMarBottom.Attribute(XName.Get("w", DocX.w.NamespaceName));

				// If w is null, this cell contains no width information.
				if (w == null)
					return double.NaN;

				// If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
				double bottomMarginInWordUnits;
				if (!double.TryParse(w.Value, out bottomMarginInWordUnits))
				{
					w.Remove();
					return double.NaN;
				}

				// 15 "word units" is equal to one pixel.
				return (bottomMarginInWordUnits / 15);
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcMar (table cell margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (tcMar == null)
				{
					tcPr.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					tcMar = tcPr.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}

				/*
                 * Get the bottom (table cell bottom margin) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcMarBottom = tcMar.Element(XName.Get("bottom", DocX.w.NamespaceName));
				if (tcMarBottom == null)
				{
					tcMar.SetElementValue(XName.Get("bottom", DocX.w.NamespaceName), string.Empty);
					tcMarBottom = tcMar.Element(XName.Get("bottom", DocX.w.NamespaceName));
				}

				// The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
				tcMarBottom.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

				// 15 "word units" is equal to one pixel. 
				tcMarBottom.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
			}
		}

		/// <summary>
		/// Set the table cell border
		/// Added by lckuiper @ 20101117
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		///using (DocX document = DocX.Create("Test.docx"))
		///{
		///    // Insert a table into this document.
		///    Table t = document.InsertTable(3, 3);
		///
		///    // Get the center cell.
		///    Cell center = t.Rows[1].Cells[1];
		///
		///    // Create a large blue border.
		///    Border b = new Border(BorderStyle.Tcbs_single, BorderSize.seven, 0, Color.Blue);
		///
		///    // Set the center cells Top, Bottom, Left and Right Borders to b.
		///    center.SetBorder(TableCellBorderType.Top, b);
		///    center.SetBorder(TableCellBorderType.Bottom, b);
		///    center.SetBorder(TableCellBorderType.Left, b);
		///    center.SetBorder(TableCellBorderType.Right, b);
		///
		///    // Save the document.
		///    document.Save();
		///}
		/// </code>
		/// </example>
		/// <param name="borderType">Table Cell border to set</param>
		/// <param name="border">Border object to set the table cell border</param>
		public void SetBorder(TableCellBorderType borderType, Border border)
		{
			/*
             * Get the tcPr (table cell properties) element for this Cell,
             * null will be return if no such element exists.
             */
			XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (tcPr == null)
			{
				Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}

			/*
             * Get the tblBorders (table cell borders) element for this Cell,
             * null will be return if no such element exists.
             */
			XElement tcBorders = tcPr.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			if (tcBorders == null)
			{
				tcPr.SetElementValue(XName.Get("tcBorders", DocX.w.NamespaceName), string.Empty);
				tcBorders = tcPr.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			}

			/*
             * Get the 'borderType' (table cell border) element for this Cell,
             * null will be return if no such element exists.
             */
			string tcbordertype;
			switch (borderType)
			{
				case TableCellBorderType.TopLeftToBottomRight:
					tcbordertype = "tl2br";
					break;
				case TableCellBorderType.TopRightToBottomLeft:
					tcbordertype = "tr2bl";
					break;
				default:
					// enum to string
					tcbordertype = borderType.ToString();
					// only lower the first char of string (because of insideH and insideV)
					tcbordertype = tcbordertype.Substring(0, 1).ToLower() + tcbordertype.Substring(1);
					break;
			}

			XElement tcBorderType = tcBorders.Element(XName.Get(borderType.ToString(), DocX.w.NamespaceName));
			if (tcBorderType == null)
			{
				tcBorders.SetElementValue(XName.Get(tcbordertype, DocX.w.NamespaceName), string.Empty);
				tcBorderType = tcBorders.Element(XName.Get(tcbordertype, DocX.w.NamespaceName));
			}

			// get string value of border style
			string borderstyle = border.Tcbs.ToString().Substring(5);
			borderstyle = borderstyle.Substring(0, 1).ToLower() + borderstyle.Substring(1);

			// The val attribute is used for the border style
			tcBorderType.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), borderstyle);

			int size;
			switch (border.Size)
			{
				case BorderSize.one: size = 2; break;
				case BorderSize.two: size = 4; break;
				case BorderSize.three: size = 6; break;
				case BorderSize.four: size = 8; break;
				case BorderSize.five: size = 12; break;
				case BorderSize.six: size = 18; break;
				case BorderSize.seven: size = 24; break;
				case BorderSize.eight: size = 36; break;
				case BorderSize.nine: size = 48; break;
				default: size = 2; break;
			}

			// The sz attribute is used for the border size
			tcBorderType.SetAttributeValue(XName.Get("sz", DocX.w.NamespaceName), (size).ToString());

			// The space attribute is used for the cell spacing (probably '0')
			tcBorderType.SetAttributeValue(XName.Get("space", DocX.w.NamespaceName), (border.Space).ToString());

			// The color attribute is used for the border color
			tcBorderType.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), border.Color.ToHex());
		}


		/// <summary>
		/// Get a table cell border
		/// Added by lckuiper @ 20101117
		/// </summary>
		/// <param name="borderType">The table cell border to get</param>
		public Border GetBorder(TableCellBorderType borderType)
		{
			// instance with default border values
			Border b = new Border();

			/*
             * Get the tcPr (table cell properties) element for this Cell,
             * null will be return if no such element exists.
             */
			XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (tcPr == null)
			{
				// uses default border style
			}

			/*
             * Get the tcBorders (table cell borders) element for this Cell,
             * null will be return if no such element exists.
             */
			XElement tcBorders = tcPr.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			if (tcBorders == null)
			{
				// uses default border style
			}

			/*
             * Get the 'borderType' (cell border) element for this Cell,
             * null will be return if no such element exists.
             */
			string tcbordertype;
			tcbordertype = borderType.ToString();

			switch (tcbordertype)
			{
				case "TopLeftToBottomRight":
					tcbordertype = "tl2br";
					break;
				case "TopRightToBottomLeft":
					tcbordertype = "tr2bl";
					break;
				default:
					// only lower the first char of string (because of insideH and insideV)
					tcbordertype = tcbordertype.Substring(0, 1).ToLower() + tcbordertype.Substring(1);
					break;
			}

			XElement tcBorderType = tcBorders.Element(XName.Get(tcbordertype, DocX.w.NamespaceName));
			if (tcBorderType == null)
			{
				// uses default border style
			}

			// The val attribute is used for the border style
			XAttribute val = tcBorderType.Attribute(XName.Get("val", DocX.w.NamespaceName));
			// If val is null, this cell contains no border information.
			if (val == null)
			{
				// uses default border style
			}
			else
			{
				try
				{
					string bordertype = "Tcbs_" + val.Value;
					b.Tcbs = (BorderStyle)Enum.Parse(typeof(BorderStyle), bordertype);
				}

				catch
				{
					val.Remove();
					// uses default border style
				}
			}

			// The sz attribute is used for the border size
			XAttribute sz = tcBorderType.Attribute(XName.Get("sz", DocX.w.NamespaceName));
			// If sz is null, this border contains no size information.
			if (sz == null)
			{
				// uses default border style
			}
			else
			{
				// If sz is not an int, something is wrong with this attributes value, so remove it
				int numerical_size;
				if (!int.TryParse(sz.Value, out numerical_size))
					sz.Remove();
				else
				{
					switch (numerical_size)
					{
						case 2: b.Size = BorderSize.one; break;
						case 4: b.Size = BorderSize.two; break;
						case 6: b.Size = BorderSize.three; break;
						case 8: b.Size = BorderSize.four; break;
						case 12: b.Size = BorderSize.five; break;
						case 18: b.Size = BorderSize.six; break;
						case 24: b.Size = BorderSize.seven; break;
						case 36: b.Size = BorderSize.eight; break;
						case 48: b.Size = BorderSize.nine; break;
						default: b.Size = BorderSize.one; break;
					}
				}
			}

			// The space attribute is used for the border spacing (probably '0')
			XAttribute space = tcBorderType.Attribute(XName.Get("space", DocX.w.NamespaceName));
			// If space is null, this border contains no space information.
			if (space == null)
			{
				// uses default border style
			}
			else
			{
				// If space is not an int, something is wrong with this attributes value, so remove it
				int borderspace;
				if (!int.TryParse(space.Value, out borderspace))
				{
					space.Remove();
					// uses default border style
				}
				else
				{
					b.Space = borderspace;
				}
			}

			// The color attribute is used for the border color
			XAttribute color = tcBorderType.Attribute(XName.Get("color", DocX.w.NamespaceName));
			if (color == null)
			{
				// uses default border style
			}
			else
			{
				// If color is not a Color, something is wrong with this attributes value, so remove it
				try
				{
					b.Color = ColorTranslator.FromHtml(string.Format("#{0}", color.Value));
				}
				catch
				{
					color.Remove();
					// uses default border style
				}
			}
			return b;
		}

		/// <summary>
		/// Gets or Sets the fill color of this Cell.
		/// </summary>
		/// <example>
		/// <code>
		/// // Create a new document.
		/// using (DocX document = DocX.Create("Test.docx"))
		/// {
		///    // Insert a table into this document.
		///    Table t = document.InsertTable(3, 3);
		///
		///    // Fill the first cell as Blue.
		///    t.Rows[0].Cells[0].FillColor = Color.Blue;
		///    // Fill the middle cell as Red.
		///    t.Rows[1].Cells[1].FillColor = Color.Red;
		///    // Fill the last cell as Green.
		///    t.Rows[2].Cells[2].FillColor = Color.Green;
		///
		///    // Save the document.
		///    document.Save();
		/// }
		/// </code>
		/// </example>
		public Color FillColor
		{
			get
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
					return Color.Empty;
				else
				{
					XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
					if (shd == null)
						return Color.Empty;
					else
					{
						XAttribute fill = shd.Attribute(XName.Get("fill", DocX.w.NamespaceName));
						if (fill == null)
							return Color.Empty;
						else
						{
							int argb = Int32.Parse(fill.Value.Replace("#", ""), NumberStyles.HexNumber);
							return Color.FromArgb(argb);
						}
					}
				}
			}

			set
			{
				/*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the tcW (table cell width) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
				if (shd == null)
				{
					tcPr.SetElementValue(XName.Get("shd", DocX.w.NamespaceName), string.Empty);
					shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
				}

				shd.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "clear");
				shd.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), "auto");
				shd.SetAttributeValue(XName.Get("fill", DocX.w.NamespaceName), value.ToHex());
			}
		}

		/// <summary></summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		/// <returns></returns>
		public override Table InsertTable(int rowCount, int columnCount)
		{
			Table table = base.InsertTable(rowCount, columnCount);
			table.mainPart = mainPart;
			InsertParagraph(); //Dmitchern, It is necessary to put paragraph in the end of the cell, without it MS-Word will say that the document is corrupted
							   //IMPORTANT: It will be better to check all methods that work with adding anything to cells
			return table;
		}

		/// <summary></summary>
		public TextDirection TextDirection
		{
			get
			{
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

				// If tcPr is null, this cell contains no width information.
				if (tcPr == null)
					return TextDirection.right;
				XElement textDirection = tcPr.Element(XName.Get("textDirection", DocX.w.NamespaceName));
				if (textDirection == null)
					return TextDirection.right;
				XAttribute val = textDirection.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (val == null)
					return TextDirection.right;

				// If val is not a VerticalAlign enum, something is wrong with this attributes value, so remove it and return VerticalAlignment.Center;
				try
				{
					return (TextDirection)Enum.Parse(typeof(TextDirection), val.Value, true);
				}

				catch
				{
					val.Remove();
					return TextDirection.right;
				}
			}
			set
			{
				/*
                    * Get the tcPr (table cell properties) element for this Cell,
                    * null will be return if no such element exists.
                    */
				XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (tcPr == null)
				{
					Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}

				/*
                 * Get the vAlign (table cell vertical alignment) element for this Cell,
                 * null will be return if no such element exists.
                 */
				XElement textDirection = tcPr.Element(XName.Get("textDirection", DocX.w.NamespaceName));
				if (textDirection == null)
				{
					tcPr.SetElementValue(XName.Get("textDirection", DocX.w.NamespaceName), string.Empty);
					textDirection = tcPr.Element(XName.Get("textDirection", DocX.w.NamespaceName));
				}

				// Set the VerticalAlignment in 'val'
				textDirection.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value.ToString());

			}
		}
	}

	/// <summary></summary>
	public class TableLook
	{

		/// <summary></summary>
		public bool FirstRow
		{
			get;
			set;
		}

		/// <summary></summary>
		public bool LastRow
		{
			get;
			set;
		}

		/// <summary></summary>
		public bool FirstColumn
		{
			get;
			set;
		}

		/// <summary></summary>
		public bool LastColumn
		{
			get;
			set;
		}

		/// <summary></summary>
		public bool NoHorizontalBanding
		{
			get;
			set;
		}

		/// <summary></summary>
		public bool NoVerticalBanding
		{
			get;
			set;
		}

	}

	/// <summary>Represents a table of contents in the document</summary>
	public class TableOfContents : DocXElement
	{
		#region TocBaseValues

		private const string HeaderStyle = "TOCHeading";
		private const int RightTabPos = 9350;
		#endregion

		private TableOfContents(DocX document, XElement xml, string headerStyle) : base(document, xml)
		{
			AssureUpdateField(document);
			AssureStyles(document, headerStyle);
		}

		internal static TableOfContents CreateTableOfContents(DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = null)
		{
			var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocXmlBase, headerStyle ?? HeaderStyle, title, rightTabPos ?? RightTabPos, BuildSwitchString(switches, lastIncludeLevel))));
			var xml = XElement.Load(reader);
			return new TableOfContents(document, xml, headerStyle);
		}

		private void AssureUpdateField(DocX document)
		{
			if (document.settings.Descendants().Any(x => x.Name.Equals(DocX.w + "updateFields"))) return;

			var element = new XElement(XName.Get("updateFields", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", true));
			document.settings.Root.Add(element);
		}

		private void AssureStyles(DocX document, string headerStyle)
		{
			if (!HasStyle(document, headerStyle, "paragraph"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocHeadingStyleBase, headerStyle ?? HeaderStyle)));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
			if (!HasStyle(document, "TOC1", "paragraph"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC1", "toc 1")));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
			if (!HasStyle(document, "TOC2", "paragraph"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC2", "toc 2")));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
			if (!HasStyle(document, "TOC3", "paragraph"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC3", "toc 3")));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
			if (!HasStyle(document, "TOC4", "paragraph"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC4", "toc 4")));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
			if (!HasStyle(document, "Hyperlink", "character"))
			{
				var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocHyperLinkStyleBase)));
				var xml = XElement.Load(reader);
				document.styles.Root.Add(xml);
			}
		}

		private bool HasStyle(DocX document, string value, string type)
		{
			return document.styles.Descendants().Any(x => x.Name.Equals(DocX.w + "style") && (x.Attribute(DocX.w + "type") == null || x.Attribute(DocX.w + "type").Value.Equals(type)) && x.Attribute(DocX.w + "styleId") != null && x.Attribute(DocX.w + "styleId").Value.Equals(value));
		}

		private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
		{
			var allSwitches = Enum.GetValues(typeof(TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
			var switchString = "TOC";
			foreach (var s in allSwitches.Where(s => s != TableOfContentsSwitches.None && switches.HasFlag(s)))
			{
				switchString += " " + s.EnumDescription();
				if (s == TableOfContentsSwitches.O)
				{
					switchString += string.Format(" '{0}-{1}'", 1, lastIncludeLevel);
				}
			}

			return switchString;
		}

	}

	#endregion

	#endregion

}

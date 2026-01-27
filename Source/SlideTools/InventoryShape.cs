using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ShapeCrawler;

//	TODO: !1 - Stopped here...
//	TODO: Working on Shape Inventory.

namespace SlideTools
{
	//*-------------------------------------------------------------------------*
	//*	InventoryShapeCollection																								*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of InventoryShapeItem Items.
	/// </summary>
	public class InventoryShapeCollection : List<InventoryShapeItem>
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
	//*	InventoryShapeItem																											*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Information about an individual shape in inventory.
	/// </summary>
	public class InventoryShapeItem
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
		//*-----------------------------------------------------------------------*
		//*	Presentation																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Presentation">Presentation</see>.
		/// </summary>
		private ShapeCrawler.Presentation mPresentation = null;
		/// <summary>
		/// Get/Set a reference to the presentation.
		/// </summary>
		public ShapeCrawler.Presentation Presentation
		{
			get { return mPresentation; }
			set { mPresentation = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Shape																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Shape">Shape</see>.
		/// </summary>
		private IShape mShape = null;
		/// <summary>
		/// Get/Set a reference to the shape.
		/// </summary>
		public IShape Shape
		{
			get { return mShape; }
			set { mShape = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Slide																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Slide">Slide</see>.
		/// </summary>
		private IUserSlide mSlide = null;
		/// <summary>
		/// Get/Set a reference to the slide upon which this shape is found.
		/// </summary>
		public IUserSlide Slide
		{
			get { return mSlide; }
			set { mSlide = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*


}

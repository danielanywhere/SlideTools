using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;

namespace SlideTools
{
	//*-------------------------------------------------------------------------*
	//*	FilterElementCollection																									*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of FilterElementItem Items.
	/// </summary>
	public class FilterElementCollection : List<FilterElementItem>
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
	//*	FilterElementItem																												*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Information about an individual filter element.
	/// </summary>
	public class FilterElementItem
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
		//*	Operator																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Operator">Operator</see>.
		/// </summary>
		private string mOperator = "=";
		/// <summary>
		/// Get/Set the value of the comparison operator to apply.
		/// </summary>
		/// <remarks>
		/// The default value of this property is '='.
		/// </remarks>
		[JsonProperty(Order = 1)]
		public string Operator
		{
			get { return mOperator; }
			set { mOperator = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Property																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Property">Property</see>.
		/// </summary>
		private string mProperty = "";
		/// <summary>
		/// Get/Set the name of the property to match.
		/// </summary>
		[JsonProperty(Order = 0)]
		public string Property
		{
			get { return mProperty; }
			set { mProperty = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Value																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Value">Value</see>.
		/// </summary>
		private string mValue = "";
		/// <summary>
		/// Get/Set the value to match on the property.
		/// </summary>
		[JsonProperty(Order = 2)]
		public string Value
		{
			get { return mValue; }
			set { mValue = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}

/*
 * Erstellt mit SharpDevelop.
 * Benutzer: mri0
 * Datum: 09.10.2009
 * Zeit: 10:32
 * 
 * Sie können diese Vorlage unter Extras > Optionen > Codeerstellung > Standardheader ändern.
 */

using System;
using System.Collections.Generic;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;

namespace MXDReader
{
	/// <summary>
	/// Description of LayerInfo.
	/// </summary>
	public class LayerInfo
	{
		private IMapDocument mxd;
		private string mxdname = "";
		private string username = "";
		private string server = "";
		private string instance = "";
		private string version = "";	
		private string name = "";
		private string owner = "";
		private string tablename = "";
		private string type = "";
		private string defquery = "";
		private string joininfo = "";
		private string minscale = "";
		private string maxscale = "";
		private string symbfields = "";
		private string parent = "";
		private string label = "";
		private string relateinfo = "";
		List<ILayer> groupLayerList;
		
		public LayerInfo(IMapDocument mapDoc)
		{
			/* Diese Eigenschaften sind für jeden Layer identisch */
			mxd = mapDoc;
			mxdname = mxd.DocumentFilename;

			UID pID = new UIDClass();
			pID.Value = "{EDAD6644-1810-11D1-86AE-0000F8751720}";
			groupLayerList = getLayerList(pID);
		}
		
		private void fillLayerProps(ILayer lyr)
		{
			name = lyr.Name;
			minscale = lyr.MinimumScale.ToString();
			maxscale = lyr.MaximumScale.ToString();
			getParentLayerNames(lyr);
			this.parent = this.parent.TrimStart('/');
		}
		
		private void getParentLayerNames(ILayer lyr)
		{
			foreach (ILayer grpLyr in groupLayerList)
			{
//				Zuerst Cast nach GroupLayer, erst dann nach ICompositeLayer
				IGroupLayer grpLyr2 = (IGroupLayer)grpLyr;
				ICompositeLayer cmpLyr = (ICompositeLayer)grpLyr2;
				for (int cmpLyrIndex=0; cmpLyrIndex < cmpLyr.Count; cmpLyrIndex++)
				{
					ILayer childLyr = cmpLyr.get_Layer(cmpLyrIndex);
					if (childLyr.Equals(lyr))
					{
						this.parent = "/" + grpLyr.Name + this.parent;
						getParentLayerNames(grpLyr);
					}
				}
			}
		}
		
		private void fillDatalayerProps(ILayer lyr)
		{
			IDataLayer dlyr = lyr as IDataLayer;
			IDatasetName dataName = dlyr.DataSourceName as IDatasetName;
			// nur für SDE-Layer werden die Connection-Infos ausgegeben
			// bei übrigen Layern wird nur der Pfad ausgegeben
			if (dataName.WorkspaceName.Type == esriWorkspaceType.esriRemoteDatabaseWorkspace)
			{
				IPropertySet props = dataName.WorkspaceName.ConnectionProperties;
				username = props.GetProperty("user").ToString();
				server = props.GetProperty("server").ToString();
				instance = props.GetProperty("instance").ToString();
				version = props.GetProperty("version").ToString();
				string[] s = dataName.Name.Split('.');
				owner = s[0];
				tablename = s[1];
			} else {
				tablename = dataName.Name;
				server = dataName.WorkspaceName.PathName;
				type = "kein SDE-Layer";
			}
		}
		
		private void fillDefQueryProps(IFeatureLayer flyr)
		{
			IFeatureLayerDefinition flyrdef = flyr as IFeatureLayerDefinition;
			if (flyrdef != null)
			{
				if (flyrdef.DefinitionExpression != "")
				{
					defquery = flyrdef.DefinitionExpression;
				}
			}
		}
		
		public void processFeatureLayer(ILayer lyr)
		{
			type = "Vektor";
			fillLayerProps(lyr);
			fillDatalayerProps(lyr);
			IFeatureLayer flyr = lyr as IFeatureLayer;
			fillDefQueryProps(flyr);
			processJoins(flyr);
			processRelates(flyr);
			
			IGeoFeatureLayer gflyr = flyr as IGeoFeatureLayer;
			if (gflyr.DisplayAnnotation == true)
			{
				IAnnotateLayerPropertiesCollection labelPropsColl = gflyr.AnnotationProperties;
				for (int collIndex=0; collIndex < labelPropsColl.Count; collIndex++)
				{
					IAnnotateLayerProperties annoLayerProps;
					IElementCollection elCol1;
					IElementCollection elCol2;
					labelPropsColl.QueryItem(collIndex, out annoLayerProps, out elCol1, out elCol2);
					string sql = annoLayerProps.WhereClause;
					ILabelEngineLayerProperties2 labelEngineProps = (ILabelEngineLayerProperties2)annoLayerProps;
					string expr = labelEngineProps.Expression;
					this.label = this.label + sql + "?" + expr + "/";
				}
			}
			this.label = this.label.TrimEnd('/');
			
			IFeatureRenderer rend = gflyr.Renderer;
			if (rend is IUniqueValueRenderer)
			{
				string felder = "";
				IUniqueValueRenderer u = rend as IUniqueValueRenderer;
				for (int i = 0; i < u.FieldCount; i++) 
				{
					felder = felder + u.get_Field(i) + "/";
				}
				symbfields = felder.TrimEnd('/') + " (UniqueValueRenderer)";
			} else if (rend is IProportionalSymbolRenderer)
			{
				IProportionalSymbolRenderer prop = rend as IProportionalSymbolRenderer;
				symbfields = prop.Field + " (ProportionalSymbolRenderer)";
			} else if (rend is IClassBreaksRenderer)
			{
				IClassBreaksRenderer cl = rend as IClassBreaksRenderer;
				symbfields = cl.Field + " (ClassBreaksRenderer)";;
			} else if (rend is ISimpleRenderer)
			{
				symbfields = "kein Feld (SimpleRenderer)";
			} else
			{
				symbfields = "unbekannter Renderer";
			}
		}
		
		private void processJoins(IFeatureLayer flyr)
		{
			IDisplayTable dispTbl = flyr as IDisplayTable;
			ITable tbl = dispTbl.DisplayTable;
			IRelQueryTable rqt;
			ITable destTable;
			IDataset dataset;
			string destName;
			string destServer;
			string destInstance;
			string destUser;
			string res = "";
			string joinType;
			// Holt iterativ alle Joins!
			while (tbl is IRelQueryTable)
			{
				rqt = (IRelQueryTable)tbl;
				IRelQueryTableInfo rqtInfo = (IRelQueryTableInfo)rqt;
				IRelationshipClass relClass = rqt.RelationshipClass;
				destTable = rqt.DestinationTable;
				if (rqtInfo.JoinType == esriJoinType.esriLeftInnerJoin)
				{
					joinType = "esriLeftInnerJoin";
				} else
				{
					joinType = "esriLeftOuterJoin";
				}
				dataset = (IDataset)destTable;
				destName = dataset.Name;
				destServer = dataset.Workspace.ConnectionProperties.GetProperty("server").ToString();
				destInstance = dataset.Workspace.ConnectionProperties.GetProperty("instance").ToString();
				destUser = dataset.Workspace.ConnectionProperties.GetProperty("user").ToString();
				res = res + "(" + destName + "/" + destServer + "/" + destInstance + "/" + destUser + "/" + relClass.OriginPrimaryKey + "/" + relClass.OriginForeignKey + "/" + joinType + ")";
				tbl = rqt.SourceTable;
			}
			joininfo = res;
		}
		
		private void processRelates(IFeatureLayer flyr) 
		{
			string res = "";
			string destName = "";
			string destServer = "";
			string destInstance = "";
			string destUser = "";
			IRelationshipClassCollection relClassColl = (IRelationshipClassCollection)flyr;
			IEnumRelationshipClass enumRelClass = relClassColl.RelationshipClasses;
			enumRelClass.Reset();
			IRelationshipClass relClass = enumRelClass.Next();
			while (relClass != null)
			{
				IDataset dset = (IDataset)relClass;
				ITable destTable = (ITable)relClass.DestinationClass;
				IDataset dsetDest = (IDataset)destTable;
				destName = dsetDest.Name;
				destServer = dsetDest.Workspace.ConnectionProperties.GetProperty("server").ToString();
				destInstance = dsetDest.Workspace.ConnectionProperties.GetProperty("instance").ToString();
				destUser = dsetDest.Workspace.ConnectionProperties.GetProperty("user").ToString();
				res = res + "(" + destName + "/" + destServer + "/" + destInstance + "/" + destUser + "/" + relClass.OriginPrimaryKey + "/" + relClass.OriginForeignKey + "/" + dset.BrowseName + ")";
				relClass = enumRelClass.Next();
			}
			
			relateinfo = res;
		}

		public void processRasterCatalogLayer(ILayer lyr)
		{
			type = "RasterKatalog";
			fillLayerProps(lyr);
			fillDatalayerProps(lyr);
			IFeatureLayer flyr = lyr as IFeatureLayer;
			fillDefQueryProps(flyr);
		}
		
		public void processAnnotationLayer(ILayer lyr)
		{
			type = "Annotation";
			fillLayerProps(lyr);
			fillDatalayerProps(lyr);
			IFeatureLayer flyr = lyr as IFeatureLayer;
			fillDefQueryProps(flyr);
		}
		
		public void processAnnotationSubLayer(ILayer lyr)
		{
			type = "AnnotationClass";
			fillLayerProps(lyr);
			// bei AnnotationSubLayern hat nur der Parent-AnnotationLayer die
			// gewünschten Informationen.
			IAnnotationSublayer sub = lyr as IAnnotationSublayer;
			ILayer parentLayer = sub.Parent as ILayer;
			fillDatalayerProps(parentLayer);
			IFeatureLayer flyr = parentLayer as IFeatureLayer;
			fillDefQueryProps(flyr);
		}

		public void processRasterLayer(ILayer lyr)
		{
			type = "Raster";
			fillLayerProps(lyr);
			fillDatalayerProps(lyr);
		}
		
		public void processGroupLayer(ILayer lyr)
		{
			type = "GroupLayer";
			fillLayerProps(lyr);			
		}
		
		public void processOtherLayer(ILayer lyr)
		{
			type = "anderer Layer";
			fillLayerProps(lyr);	
		}
		
		public string writeCSV()
		{
			string output = mxdname + ";" + name + ";" + type + ";" + owner + ";" + tablename + ";" + server + ";" + instance + ";" + username.ToUpper() + ";" + version + ";" + minscale + ";" + maxscale + ";" + defquery + ";" + joininfo + ";" + symbfields + ";" + parent + ";" + label + ";" + relateinfo;
			return output;
		}
		
		private List<ILayer> getLayerList(UID id)
		{
			List<ILayer> res = new List<ILayer>();
			// try-catch ist nötig, da get_Layers() eine Exception auslöst,
			// wenn es keine passenden Layer findet!
			try
			{
				IEnumLayer lyrs = mxd.get_Map(0).get_Layers(id, true);
				lyrs.Reset();
				ILayer lyr = lyrs.Next();
				while (lyr != null)
				{
					res.Add(lyr);
					lyr = lyrs.Next();
				}
			}
			catch 
			{
			}
			return res;
		}
	}
}

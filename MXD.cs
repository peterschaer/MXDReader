/*
 * Erstellt mit SharpDevelop.
 * Benutzer: mri0
 * Datum: 09.10.2009
 * Zeit: 10:33
 * 
 * Sie können diese Vorlage unter Extras > Optionen > Codeerstellung > Standardheader ändern.
 */

using System;
using System.Collections.Generic;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;

namespace MXDReader
{
	/// <summary>
	/// Description of MXD.
	/// </summary>
	public class MXD
	{
		string path;
		IMapDocument mapDoc;
		List<LayerInfo> lyrInfos = new List<LayerInfo>();
		List<ILayer> layerList;
		List<ILayer> processedLayers = new List<ILayer>();
		
		AoInitialize init = new AoInitializeClass();
		string[] LayerTypeUIDs = new string[7];
		enum LayerTypes
		{
			GroupLayer = 0,
			AnnotationLayer = 1,
			AnnotationSubLayer = 2,
			RasterLayer = 3,
			RasterCatalogLayer = 4,
			FeatureLayer = 5,
			OtherLayer = 6
		}
		
		
		public MXD(string mxdPath)
		{
			path = mxdPath;
			// GUIDs finden: http://support.esri.com/index.cfm?fa=knowledgebase.techarticles.articleShow&d=31115
			LayerTypeUIDs[0] = "{EDAD6644-1810-11D1-86AE-0000F8751720}";
			LayerTypeUIDs[1] = "{5CEAE408-4C0A-437F-9DB3-054D83919850}";
			LayerTypeUIDs[2] = "{DBCA59AC-6771-4408-8F48-C7D53389440C}";
			LayerTypeUIDs[3] = "{D02371C7-35F7-11D2-B1F2-00C04F8EDEFF}";
			LayerTypeUIDs[4] = "{605BC37A-15E9-40A0-90FB-DE4CC376838C}";
			LayerTypeUIDs[5] = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}";
			LayerTypeUIDs[6] = "";
			
			if (InitializeLicense() == esriLicenseStatus.esriLicenseCheckedOut)
			{
				mapDoc = new MapDocumentClass();
				if (mapDoc.get_IsPresent(path))
				{
					if (mapDoc.get_IsMapDocument(path))
					{
						mapDoc.Open(path, null);
						if (mapDoc.DocumentType == esriMapDocumentType.esriMapDocumentTypeMxd)
						{
							activateMapDocument();
							
							// alle Layer holen
							layerList = getLayerList(null);
							
							// alle GroupLayer holen
							processLayer(LayerTypes.GroupLayer);
							
							// alle AnnotationLayer holen
							processLayer(LayerTypes.AnnotationLayer);
							
							// alle AnnotationSubLayer holen
							processLayer(LayerTypes.AnnotationSubLayer);
			
							// alle RasterLayer holen
							processLayer(LayerTypes.RasterLayer);
							
							// alle RasterCatalogLayer holen
							processLayer(LayerTypes.RasterCatalogLayer);
			
							// alle FeatureLayer holen
							processLayer(LayerTypes.FeatureLayer);
							
							// übrige Layer holen
							foreach (ILayer lyr in layerList)
							{
								if (!processedLayers.Contains(lyr))
								{
									LayerInfo lyrInfo = new LayerInfo(mapDoc);
									lyrInfo.processOtherLayer(lyr);
									lyrInfos.Add(lyrInfo);
									processedLayers.Add(lyr);
								}
							}
							
							// Resultate ausgeben
							foreach (LayerInfo info in lyrInfos) 
							{
								Console.WriteLine(info.writeCSV());
							}
							
							// Dokument schliessen
							mapDoc.Close();	
							
						} else {
							Console.Error.WriteLine("FEHLER: kein gültiges MXD-File");
						}
					} else {
						Console.Error.WriteLine("FEHLER: kein gültiges MXD-File");
					}
				} else {
					Console.Error.WriteLine("FEHLER: Angegebenes File nicht gefunden");
				}

				
			} else 
			{
				Console.Error.WriteLine("FEHLER: Keine ArcView-Lizenz verfügbar");
			}
			
			// Lizenz zurückgeben
			TerminateLicense();
		}
		
		private void processLayer(LayerTypes lt)
		{
			UID pID = new UIDClass();
			int layerTypeID = (int)lt;
			pID.Value = LayerTypeUIDs[(int)lt];
			List<ILayer> FoundLayerList = getLayerList(pID);
			foreach (ILayer lyr in FoundLayerList)
			{
				if (!processedLayers.Contains(lyr))
			    {
					LayerInfo lyrInfo = new LayerInfo(mapDoc);
					switch (lt)
					{
						case LayerTypes.GroupLayer:
							lyrInfo.processGroupLayer(lyr);
							break;
						case LayerTypes.AnnotationLayer:
							lyrInfo.processAnnotationLayer(lyr);
							break;
						case LayerTypes.AnnotationSubLayer:
							lyrInfo.processAnnotationSubLayer(lyr);
							break;
						case LayerTypes.RasterLayer:
							lyrInfo.processRasterLayer(lyr);
							break;
						case LayerTypes.RasterCatalogLayer:
							lyrInfo.processRasterCatalogLayer(lyr);
							break;
						case LayerTypes.FeatureLayer:
							lyrInfo.processFeatureLayer(lyr);
							break;
						case LayerTypes.OtherLayer:
							lyrInfo.processOtherLayer(lyr);
							break;
					}
					lyrInfos.Add(lyrInfo);
					processedLayers.Add(lyr);
			    }
			}
		}
		
		private void activateMapDocument()
		{
			// Das MapDocument muss gemäss ESRI Dev-Help aktiviert werden,
			// bevor es ausgelesen wird. Sonst sind die Informationen evtl. falsch
			int pointer = Win32APICall.GetDesktopWindow().ToInt32();
			// Map aktivieren
			mapDoc.ActiveView.Activate(pointer);
			// PageLayout aktivieren
			IActiveView av = mapDoc.PageLayout as IActiveView;
			av.Activate(pointer);
		}
		
		private List<ILayer> getLayerList(UID id)
		{
			List<ILayer> res = new List<ILayer>();
			// try-catch ist nötig, da get_Layers() eine Exception auslöst,
			// wenn es keine passenden Layer findet!
			try
			{
				IEnumLayer lyrs = mapDoc.get_Map(0).get_Layers(id, true);
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
		
		public string Path {
			get { return path; }
		}
		
		private esriLicenseStatus InitializeLicense()
		{
			esriLicenseStatus LicenseStatus = esriLicenseStatus.esriLicenseUnavailable;
			LicenseStatus = init.IsProductCodeAvailable(esriLicenseProductCode.esriLicenseProductCodeBasic);
			if (LicenseStatus == esriLicenseStatus.esriLicenseAvailable)
			{
				LicenseStatus = init.Initialize(esriLicenseProductCode.esriLicenseProductCodeBasic);
			} else {
				LicenseStatus = esriLicenseStatus.esriLicenseUnavailable;
			}
			return LicenseStatus;
		}
		
		private void TerminateLicense()
		{
			init.Shutdown();
		}
		
	}
}

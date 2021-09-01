using System;
using System.Diagnostics;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Resources;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Windows.Forms;
using System.IO.Packaging;
using Daisy.SaveAsDAISY.Conversion;
namespace Daisy.SaveAsDAISY.Forms
{
	public partial class ProgressValidation : Form
	{
		const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		ArrayList docValidation = new ArrayList();
		XmlDocument currentDocXml = new XmlDocument();
		XmlDocument validation_xml = new XmlDocument();
		private Exception exception;
		int count;
		private bool converting, cancel;
		int c_count = 0, j;
		XmlNode eStyle, f_siblingName, f_siblingstyle, p_sibling, p_siblingStyle, parent, parentStyle, parentFollowingSibling, parentFollowingSiblingStyle;
		XmlNode parentPreviousSibling, parentPreviousSiblingStyle, child, followingChild, childStyle, nextRule, errMessage;

		//Constructor to initialize maximum value of the progress bar 
		public ProgressValidation(XmlDocument currentDocXml, XmlDocument validation_xml)
		{
			InitializeComponent();
			//Counting total nodes of validation xml
			for (int i = 0; i < validation_xml.DocumentElement.ChildNodes.Count; i++)
			{
				count = count + validation_xml.DocumentElement.ChildNodes[i].ChildNodes.Count;
			}
			//Counting total nodes of document xml
			for (int i = 0; i < currentDocXml.DocumentElement.ChildNodes.Count; i++)
			{
                if (currentDocXml.DocumentElement.ChildNodes[i].ChildNodes.Count != 0)
                {
                    count = count*currentDocXml.DocumentElement.ChildNodes[i].ChildNodes.Count;
                }
			}
			//Assingning maximum value of the progress bar
			this.ProgressBar.Maximum = count;
			this.currentDocXml = currentDocXml;
			this.validation_xml = validation_xml;
			lblProgress.Text = "Loading validation details.. Please wait.";
		}

		//Property to get Arraylist docValidation
		public ArrayList GetFileNames
		{
			get { return docValidation; }
		}

		//Property to get updated document xml
		public XmlDocument XmlFileNames
		{
			get { return currentDocXml; }
		}

		public void AddValidationError(string error)
		{
			docValidation.Add(error);
		}

		//Function to increment profress bar
		public void DoTranslation()
		{
			try
			{
				OoxPageStylesValidator ooxPageStylesValidator = new OoxPageStylesValidator(this);

                this.TOC_BodyMatter_Validation(this.currentDocXml);

				//Looping through each node of document xml and validation xml
				for (int i = 0; i < this.currentDocXml.DocumentElement.ChildNodes.Count; i++)
				{
					for (j = 0; j < this.currentDocXml.DocumentElement.ChildNodes[i].ChildNodes.Count; j++)
					{
						ooxPageStylesValidator.ValidateNode(currentDocXml.DocumentElement.ChildNodes[i].ChildNodes[j]);

						for (int k = 0; k < this.validation_xml.DocumentElement.ChildNodes.Count; k++)
						{
							for (int l = 0; l < this.validation_xml.DocumentElement.ChildNodes[k].ChildNodes.Count; l++)
							{
								c_count = 0;
								//Checking if node of document xml matches the node of validation xml
								if (string.Compare(this.currentDocXml.DocumentElement.ChildNodes[i].ChildNodes[j].Name, "w:" + validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].FirstChild.InnerText) == 0)
								{
									//Calling function to get the node of validation xml
									RulesElements(k, l);
									//Calling function to campare document xml node with validation xml node accoring to the rule
									Compare(this.currentDocXml.DocumentElement.ChildNodes[i].ChildNodes[j], validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l]);
								}
								//Incrementing progress bar
								this.ProgressBar.Increment(1);
								float percentShow = (this.ProgressBar.Value * 100) / this.ProgressBar.Maximum;
								lblProgress.Text = "Validation Process: " + percentShow.ToString() + "% completed";
								Application.DoEvents();

								//Checking if progress bar is canceled
								if (cancel)
								{
									throw new CancelledException();
								}
							}
						}
					}
				}

				ooxPageStylesValidator.ValidateLastStyle();

				//Work complete
				WorkComplete(null);
			}
			//Catch block for work incomplete exception
			catch (Exception e)
			{
				WorkComplete(e);
			}
		}

		//Fuction to get the status of the work
		private void WorkComplete(Exception e)
		{
			//Work complete
			if (e == null)
			{
				DialogResult = DialogResult.OK;
			}
			else
			{
				//If work canceled
				if (e is CancelledException)
				{
					DialogResult = DialogResult.Cancel;
				}
				//If work aborted
				else
				{
					this.exception = e;
					DialogResult = DialogResult.Abort;
				}
			}
		}

		public class CancelledException : Exception
		{
		}

		public Exception Exception
		{
			get
			{
				return this.exception;
			}
		}

		//Function to get the node of validation xml
		public void RulesElements(int k, int l)
		{
			eStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ElementStyle");
			f_siblingName = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("FollowingSiblingName");
			f_siblingstyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("FollowingSiblingStyle");
			p_sibling = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("PrecedingSiblingName");
			p_siblingStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("PrecedingSiblingStyle");
			parent = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("Parent");
			parentStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ParentStyle");
			parentFollowingSibling = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ParentFollowingSibling");
			parentFollowingSiblingStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ParentFollowingSiblingStyle");
			parentPreviousSibling = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ParentPrecedingSibling");
			parentPreviousSiblingStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ParentPrecedingSiblingStyle");
			child = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("Child");
			followingChild = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ChildFollowingSibling");
			childStyle = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ChildStyle");
			nextRule = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("NextRule");
			errMessage = validation_xml.DocumentElement.ChildNodes[k].ChildNodes[l].SelectSingleNode("ErrMessage");
		}

		//Function to generate Id for bookmark
		public long GenerateId()
		{
			byte[] buffer = Guid.NewGuid().ToByteArray();
			return BitConverter.ToInt64(buffer, 0);
		}

        /// <summary>
        /// TOC and BODYMATTER style validation.
        /// <para>1. If document contents automatic TOC and doesn't contents BODYMATTER style after TOC, DAISY must show validation error.</para>
        /// <para>2. If document contents automatic TOC and contents BODYMATTER style before TOC, DAISY must show validation error.</para>
        /// <para>3. If document contents automatic TOC and contents BODYMATTER style after TOC, DAISY must generate TOC in BODYMATTER style without movement to the top of document (to FRONTMATTER).</para>
        /// </summary>
        /// <param name="xmlDocument"></param>
        private void TOC_BodyMatter_Validation(XmlDocument xmlDocument)
        {
            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", docNamespace);

            XmlNode tocNode = xmlDocument.DocumentElement.SelectSingleNode("//w:pPr[w:docPartObj/w:docPartGallery/@w:val='Table of Contents' or substring(w:pStyle/@w:val, 1, 3)='TOC']", nsManager);
            if (tocNode == null)
                tocNode = xmlDocument.DocumentElement.SelectSingleNode("//w:sdtPr[w:docPartObj/w:docPartGallery/@w:val='Table of Contents' or substring(w:pStyle/@w:val, 1, 3)='TOC']", nsManager);
            if (tocNode != null)
            {
                XmlNode bodymatterNode = tocNode.SelectSingleNode("following::w:pPr/w:pStyle[substring(@w:val, 1, 10)='Bodymatter']", nsManager);
				
                if (bodymatterNode == null)
                    this.CreateBookmark(" :" + Properties.Resources.TocBodymatterValidationMessage, tocNode, tocNode, tocNode);
            }
        }

		//Function to campare document xml node with validation xml node accoring to the rule
		public void Compare(XmlNode nd2, XmlNode nd1)
		{
			NameTable nt = new NameTable();
			XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
			nsManager.AddNamespace("w", docNamespace);

			//Checking validation xml <ElementStyle> element exists
			if (eStyle != null)
			{
				//Calling function to compare style of document xml current element with <ElementStyle> element of validation xml
				CheckElementStyle(nd2, nd1, nsManager);
			}

			//Checking validation xml <FollowingSiblingName> element exists
			if (f_siblingName != null)
			{
				//Calling function to compare next sibling of document xml current element with <FollowingSiblingName> element of validation xml
				CheckFollowingSibling(nd2, nd1, nsManager);
			}

			//Checking validation xml <FollowingSiblingStyle> element exists
			if (f_siblingstyle != null)
			{
				//Calling function to compare next sibling style of document xml current element with <FollowingSiblingStyle> element of validation xml
				CheckFsiblingStyle(nd2, nd1, nsManager);
			}

			//Checking validation xml <PrecedingSiblingName> element exists
			if (p_sibling != null)
			{
				//Calling function to compare previous sibling of document xml current element with <PrecedingSiblingName> element of validation xml
				CheckPrecedingSibling(nd2, nd1, nsManager);
			}

			//Checking validation xml <PrecedingSiblingStyle> element exists
			if (p_siblingStyle != null)
			{
				//Calling function to compare previous sibling style of document xml current element with <PrecedingSiblingStyle> element of validation xml
				CheckPsiblingStyle(nd2, nd1, nsManager);
			}

			//Checking validation xml <Child> element exists
			if (child != null)
			{
				//Calling function to compare child of document xml current element with <Child > element of validation xml
				CheckChild(nd2, nd1, nsManager);
			}

			//Checking validation xml <NextRule> element exists
			if (nextRule != null && c_count == nd1.ChildNodes.Count - 2)
			{
				if (nextRule.InnerText.EndsWith("-"))
				{
					if (nd2.FirstChild != null)
					{
						XmlNode compNode = CheckforNames(validation_xml.SelectSingleNode("//" + nextRule.InnerText), nd2.FirstChild);
						if (compNode != null)
							CompareChild(validation_xml.SelectSingleNode("//" + nextRule.InnerText), compNode, nsManager);
					}
				}
				//Calling function which recursiveley check for <NextRule> element of validation xml
				//For each <NextRule> element it will check if the <Child> element of validation xml matches with the child element of the current element of document xml

				else if (f_siblingName != null && nd2.NextSibling != null)
				{
					if (j == 0)
					{
						XmlNode compNode = CheckforNames(validation_xml.SelectSingleNode("//" + nextRule.InnerText), nd2.NextSibling);
						if (compNode != null)
							CompareChild(validation_xml.SelectSingleNode("//" + nextRule.InnerText), compNode, nsManager);
					}
					else if (nd2.PreviousSibling != null)
					{
						if (nd2.PreviousSibling.Name != "w:tbl" && nd2.PreviousSibling.SelectSingleNode("w:r/w:drawing", nsManager) == null)
						{
							XmlNode compNode = CheckforNames(validation_xml.SelectSingleNode("//" + nextRule.InnerText), nd2.NextSibling);
							if (compNode != null)
								CompareChild(validation_xml.SelectSingleNode("//" + nextRule.InnerText), compNode, nsManager);
						}

					}
				}
			}
			else if (child != null && nd2.HasChildNodes && nd1.Name.EndsWith("T"))
			{
				for (int ti = 0; ti < nd2.ChildNodes.Count; ti++)
				{
					if (nd2.ChildNodes[ti].Name == "w:tr")
					{
						for (int tj = 0; tj < nd2.ChildNodes[ti].ChildNodes.Count; tj++)
						{
							if (nd2.ChildNodes[ti].ChildNodes[tj].Name == "w:tc")
							{
								for (int tk = 0; tk < nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes.Count; tk++)
								{
									if (nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].Name == "w:p")
									{
										for (int tm = 0; tm < nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes.Count; tm++)
										{
											if (nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes[tm].Name == "w:pPr")
											{
												if (nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes[tm].FirstChild.Name == "w:pStyle")
												{
													if (nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes[tm].FirstChild.Attributes[0].Value.Substring(0, 7) == nd1.SelectSingleNode("//ChildStyle").InnerText)
                                                    {
                                                        this.CreateBookmark(nd1,
                                                                            nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes[tm].ParentNode.FirstChild,
                                                                            nd2.ChildNodes[ti].ChildNodes[tj].ChildNodes[tk].ChildNodes[tm].ParentNode.LastChild);
													}
												}
											}

										}
									}

								}
							}

						}
					}

				}


			}
			//If <NextRule> element is not present and all comparison is true
			else if (nextRule == null && c_count == nd1.ChildNodes.Count - 2)
            {
                this.CreateBookmark(nd1, nd2, nd2);
			}
		}

		public XmlNode CheckforNames(XmlNode nodeValidXml, XmlNode nodeDocXml)
		{
			XmlNode tempNode = null;
			if (nodeDocXml.HasChildNodes)
			{
				for (int i = 0; i < nodeDocXml.ChildNodes.Count; i++)
				{
					// For each child matching the current current rule "ElementName"
					if ("w:" + nodeValidXml.ChildNodes[0].InnerText == nodeDocXml.ChildNodes[i].Name)
					{
						tempNode = nodeDocXml.ChildNodes[i]; // store for return value
						if (nodeValidXml.SelectSingleNode("//NextRule") != null) // if the current has a "NextRule" element
						{
							// Select the "NextRule" node
							// This test does not make any sense, as the "next rule" result in a text node and not a rule node
							if (NextNodePresent(nodeValidXml.SelectSingleNode("//NextRule"), tempNode))
								break;
						}
					}
				}
			}
			return tempNode;
		}
		/// <summary>
		/// Test if the document node contains a node matching the first child value of the validation rule node
		/// (allegedly 
		/// </summary>
		/// <param name="nodeValidXml"></param>
		/// <param name="nodeDocXml"></param>
		/// <returns></returns>
		public bool NextNodePresent(XmlNode nodeValidXml, XmlNode nodeDocXml)
		{
			if (nodeDocXml.SelectSingleNode("//" + nodeValidXml.ChildNodes[0].InnerText) != null)
				return true;
			else
				return false;
		}

		//Function which recursiveley check for <NextRule> element of validation xml
		//For each <NextRule> element it will check if the <Child> element of validation xml matches with the child element of the current element of document xml
		public void CompareChild(XmlNode nd1, XmlNode nd2, XmlNamespaceManager nsManager)
		{
			String name = nd1.SelectSingleNode("//" + nd1.Name + "/ElementName").InnerText;
			if (nd2.SelectSingleNode("//" + "w:" + name, nsManager) != null)
			{

				if ("w:" + name == nd2.SelectSingleNode("//" + "w:" + name, nsManager).Name)
				{
					XmlNode tempNode = nd1.SelectSingleNode("//" + nd1.Name + "/NextRule");
					XmlNode elementStyle = nd1.SelectSingleNode("//" + nd1.Name + "/ElementStyle");
					if (tempNode != null)
					{
						if (tempNode.InnerText.EndsWith("-") && tempNode.SelectSingleNode("//FollowingSibling").Name == "FollowingSibling")
						{
							if (nd2.NextSibling != null)
							{
								XmlNode compNode = CheckforNames(validation_xml.SelectSingleNode("//" + tempNode.InnerText), nd2.NextSibling);
								if (compNode != null)
									CompareChild(validation_xml.SelectSingleNode("//" + tempNode.InnerText), compNode, nsManager);
							}
						}

						else
						{
							if (nd2 != null)
							{
								XmlNode compNode = CheckforNames(validation_xml.SelectSingleNode("//" + tempNode.InnerText), nd2);
								if (compNode != null)
									CompareChild(validation_xml.SelectSingleNode("//" + tempNode.InnerText), compNode, nsManager);
							}
						}
					}
					else if (elementStyle != null)
					{
						if (elementStyle.InnerText == nd2.Attributes[0].Value)
						{
                            this.CreateBookmark(nd1,
                                                nd2.ParentNode.ParentNode.ParentNode,
                                                nd2.ParentNode.ParentNode.ParentNode);
						}
					}
					else
					{
                        this.CreateBookmark(nd1,
                                            nd2.ParentNode.ParentNode.PreviousSibling,
                                            nd2.ParentNode.ParentNode.PreviousSibling);
					}
				}
			}
		}

        private void CreateBookmark(string errorMessage, XmlNode nd1, XmlNode prependingNode, XmlNode appendingNode)
        {
            //Creating boobkmarkstart element
            XmlElement bookStart = currentDocXml.CreateElement("w:bookmarkStart", docNamespace);
            String nameValidation = nd1.Name + GenerateId();
            String Id = GenerateId().ToString();
            //Creating attribute id for bookmarkstart
            bookStart.SetAttribute("id", docNamespace, Id);
            //Creating attribute name for bookmark
            bookStart.SetAttribute("name", docNamespace, nameValidation);
            //Creating boobkmarkend element
            XmlElement bookEnd = currentDocXml.CreateElement("w:bookmarkEnd", docNamespace);
            //Creating attribute id for bookmarkend
            bookEnd.SetAttribute("id", docNamespace, Id);
            //Adding error message to arraylist docValidation
            docValidation.Add(errorMessage + "|" + nameValidation);
            //Prepending bookmarkstart tag
            prependingNode.PrependChild(bookStart);
            //Appending bookmarkend tag
            appendingNode.AppendChild(bookEnd);
        }

        private void CreateBookmark(XmlNode nd1, XmlNode prependingNode, XmlNode appendingNode)
        {
            this.CreateBookmark(nd1.SelectSingleNode("//" + nd1.Name + "/ErrMessage").InnerText, nd1, prependingNode, appendingNode);
        }

	    //Function to compare style of document xml current element with <ElementStyle> element of validation xml
		public void CheckElementStyle(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.SelectSingleNode("w:pPr/w:pStyle", nsManager) != null)
			{
				if (nd1.SelectSingleNode("//" + nd1.Name + "/ElementStyle").InnerText == nd2.SelectSingleNode("w:pPr/w:pStyle", nsManager).Attributes[0].Value)
				{
					c_count++;
				}
			}
		}

		//Function to compare next sibling of document xml current element with <FollowingSiblingName> element of validation xml
		public void CheckFollowingSibling(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.NextSibling != null)
			{
				if ("w:" + nd1.SelectSingleNode("//" + nd1.Name + "/FollowingSiblingName").InnerText == nd2.NextSibling.Name)
				{
					c_count++;
				}
			}
		}

		//Function to compare next sibling style of document xml current element with <FollowingSiblingStyle> element of validation xml
		public void CheckFsiblingStyle(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.NextSibling.SelectSingleNode("w:pPr/w:pStyle", nsManager) != null)
			{
				if (nd1.SelectSingleNode("//" + nd1.Name + "/FollowingSiblingStyle").InnerText == nd2.NextSibling.SelectSingleNode("w:pPr/w:pStyle", nsManager).Attributes[0].Value)
				{
					c_count++;
				}
			}
		}

		public void CheckFsiblingChild(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.HasChildNodes)
			{
				String name = nd1.SelectSingleNode("//" + nd1.Name + "/Child").InnerText;
				if (nd2.SelectSingleNode("//" + "w:" + name, nsManager) != null)
				{

					if ("w:" + name == nd2.SelectSingleNode("//" + "w:" + name, nsManager).NextSibling.Name)
					{
						c_count++;
					}
				}
			}
		}

		//Function to compare previous sibling of document xml current element with <PrecedingSiblingName> element of validation xml
		public void CheckPrecedingSibling(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (j != 0)
			{
				if (nd2.PreviousSibling != null)
				{
					if ("w:" + nd1.SelectSingleNode("//" + nd1.Name + "/PrecedingSiblingName").InnerText == nd2.PreviousSibling.Name)
					{
						c_count++;
					}
				}
			}
		}

		//Function to compare previous sibling style of document xml current element with <PrecedingSiblingStyle> element of validation xml
		public void CheckPsiblingStyle(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (j != 0)
			{
				if (nd2.PreviousSibling.SelectSingleNode("w:pPr/w:pStyle", nsManager) != null)
				{
					if (nd1.SelectSingleNode("//" + nd1.Name + "/PrecedingSiblingStyle").InnerText == nd2.PreviousSibling.SelectSingleNode("w:pPr/w:pStyle", nsManager).Attributes[0].Value)
					{
						c_count++;
					}
				}
			}
		}

		//Function to compare child of document xml current element with <Child > element of validation xml 
		public void CheckChild(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.HasChildNodes)
			{
				String name = nd1.SelectSingleNode("//" + nd1.Name + "/Child").InnerText;
                if (nd2.SelectSingleNode("w:pPr/" + "w:" + name, nsManager) != null)
				{
                    if ("w:" + name == nd2.SelectSingleNode("w:pPr/" + "w:" + name, nsManager).Name)
					{
						c_count++;
					}
				}
			}
		}

		public void CheckChildStyle(XmlNode nd2, XmlNode nd1, XmlNamespaceManager nsManager)
		{
			if (nd2.HasChildNodes)
			{
				String name = nd1.SelectSingleNode("//" + nd1.Name + "/ChildStyle").InnerText;
				// String nameSty = nd1.SelectSingleNode("//" + nd1.Name + "/ChildStyle").InnerText;
				if (nd2.SelectSingleNode("w:r/w:rPr/w:rStyle", nsManager) != null)
				{
					if ("w:" + name == nd2.NextSibling.SelectSingleNode("w:r/w:rPr/w:rStyle", nsManager).Attributes[0].Value)
					{
						c_count++;
					}
				}
			}
		}

		//Event progress bar activated
		private void ProgressValidation_Activated(object sender, EventArgs e)
		{
			if (!converting)
			{
				converting = true;
				Application.DoEvents();
				DoTranslation();
				converting = false;
			}
		}

		//Function for checking if progress bar is canceled
		private void button1_Click(object sender, EventArgs e)
		{
			cancel = true;
		}
	}
}
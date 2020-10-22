using GemBox.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PowerPointGenerator
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private const string imageSearchUrl = "https://www.google.com/search?q={{QUERY_STRING}}&tbm=isch&sclient=img";
		private Regex GoogleImageRegEx;
		private List<Image> Images;
		private Dictionary<Image, bool> Selection;

		static MainWindow()
		{
			ComponentInfo.SetLicense( "FREE-LIMITED-KEY" );
		}

		public MainWindow()
		{
			InitializeComponent();
			GoogleImageRegEx = new Regex( @"https:\/\/encrypted-tbn0\.gstatic\.com\/images\?q=tbn:[\w&-]*;s" );
			Images = new List<Image>();
			Selection = new Dictionary<Image, bool>();
		}

		private List<string> GetBoldedText( RichTextBox box )
		{
			List<string> allWords = new List<string>();

			foreach( Paragraph p in SearchText.Document.Blocks )
			{
				foreach( var inline in p.Inlines )
				{
					if( inline.FontWeight == FontWeights.Bold )
					{
						allWords.Add( new System.Windows.Documents.TextRange( inline.ContentStart, inline.ContentEnd ).Text );
					}
				}
			}

			return allWords;
		}

		private void Search_Click( object sender, RoutedEventArgs e )
		{
			List<string> allWords = new List<string>();
			Image_Panel.Children.Clear();
			Selection.Clear();
			allWords.Add( Title.Text );

			allWords.AddRange( GetBoldedText( SearchText ) );

			var queryString = imageSearchUrl.Replace( "{{QUERY_STRING}}", HttpUtility.UrlEncode( string.Join( " ", allWords ) ) );
			GetImages( queryString );
		}

		private void Bold_Click( object sender, RoutedEventArgs e )
        {
			if(SearchText.Selection.GetPropertyValue( System.Windows.Documents.TextElement.FontWeightProperty).Equals(FontWeights.Bold))
            {
				SearchText.Selection.ApplyPropertyValue( System.Windows.Documents.TextElement.FontWeightProperty, FontWeights.Normal );
            }
            else
            {
				SearchText.Selection.ApplyPropertyValue( System.Windows.Documents.TextElement.FontWeightProperty, FontWeights.Bold );
            }
        }

		private void Select_Click( object sender, RoutedEventArgs e )
		{
			//in one spot it says minimum of 3 images on slide, but final point in keeping from passing says UP TO 3 images ¯\_(ツ)_/¯
			if( Selection.Where( x => x.Value ).Count() != 3 )
			{
				return;
			}

			var presentation = new PresentationDocument();
			var slide = presentation.Slides.AddNew( SlideLayoutType.Custom );
			var textBox = slide.Content.AddTextBox( ShapeGeometryType.Rectangle, 10, 2, 5, 4, LengthUnit.Centimeter );

			var title = textBox.AddParagraph();

			title.AddRun( Title.Text );


			var bodyTextBox = slide.Content.AddTextBox( ShapeGeometryType.Rectangle, 20, 2, 5, 4, LengthUnit.Centimeter );

			var body = bodyTextBox.AddParagraph();

			var bodyText = new System.Windows.Documents.TextRange( SearchText.Document.ContentStart, SearchText.Document.ContentEnd ).Text.Replace( "\r", " " );
			bodyText = bodyText.Replace( "\n", " " );

			body.AddRun( bodyText );

			double width = 0.0;
			foreach( var image in Selection )
			{
				if( image.Value )
				{
					slide.Content.AddPicture( image.Key.Source.ToString(), width, 250, image.Key.Width, image.Key.Height );
					width += image.Key.Width;
				}
			}

			presentation.Save( $"Results_{Guid.NewGuid()}.pptx" );
			Image_Panel.Children.Clear();
			Images.Clear();
			Selection.Clear();
		}

		private Image CreateImage( string source, double width = 125.0, double height = 125.0 )
		{
			var image = new Image();

			BitmapImage bitmap = new BitmapImage();
			bitmap.BeginInit();
			bitmap.UriSource = new Uri( source, UriKind.Absolute );
			bitmap.EndInit();

			image.Source = bitmap;
			image.Width = width;
			image.Height = height;
			image.Opacity = 0.5;

			return image;
		}

		private void SelectImage( object sender, MouseButtonEventArgs e )
		{
			var count = Selection.Where( x => x.Value ).Count();

			Image image = (Image)sender;
			if( !Selection.ContainsKey( image ) )
			{
				if( count < 3 )
				{
					Selection[ image ] = true;
					image.Opacity = 1.0;
				}
			}
			else
			{
				if( Selection[ image ] )
				{
					image.Opacity = 1.0;
					Selection[ image ] = false;
				}
				else if( count < 3 )
				{
					image.Opacity = 0.5;
					Selection[ image ] = true;
				}
			}
		}


		private async void GetImages( string url )
		{
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create( url );
			var response = (HttpWebResponse)await request.GetResponseAsync();

			if( response.StatusCode == HttpStatusCode.OK )
			{
				Stream receiveStream = response.GetResponseStream();
				StreamReader readStream;

				if( string.IsNullOrWhiteSpace( response.CharacterSet ) )
				{
					readStream = new StreamReader( receiveStream );
				}
				else
				{
					readStream = new StreamReader( receiveStream, Encoding.GetEncoding( response.CharacterSet ) );
				}


				var lines = ( await readStream.ReadToEndAsync() ).Split( new char[] { '\n' } );
				for( int j = 0; j < lines.Length; j++ )
				{
					var line = lines[ j ];
					var matches = GoogleImageRegEx.Matches( line );

					for( int i = 0; i < matches.Count; i++ )
					{
						var image = CreateImage( matches[ i ].Value );
						image.MouseDown += SelectImage;
						Image_Panel.Children.Add( image );
						Images.Add( image );
					}
				}


				response.Close();
				readStream.Close();
			}
		}

	}
}

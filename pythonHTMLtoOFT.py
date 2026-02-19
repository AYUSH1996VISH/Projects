"""
How to run - python pythonHTMLtoOFT.py --gui
HTML to Outlook OFT Converter - Enterprise Edition with Preview
Version: 3.0.0 - Full Template Preview System
Author: Advanced Email Engineering Team
"""

import os
import sys
import re
import base64
import logging
import mimetypes
import tempfile
import shutil
import ssl
import urllib.request
import webbrowser
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from datetime import datetime
from urllib.parse import urlparse, urljoin
import hashlib

# Third-party imports with fallback handling
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    print("Warning: pywin32 not installed. Install with: pip install pywin32")

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False
    print("Warning: BeautifulSoup not installed. Install with: pip install beautifulsoup4")

try:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    print("Warning: requests not installed. Install with: pip install requests")

try:
    import cssutils
    cssutils.log.setLevel(logging.CRITICAL)
    CSSUTILS_AVAILABLE = True
except ImportError:
    CSSUTILS_AVAILABLE = False
    print("Warning: cssutils not installed. Install with: pip install cssutils")

try:
    from premailer import Premailer
    PREMAILER_AVAILABLE = True
except ImportError:
    PREMAILER_AVAILABLE = False
    print("Warning: premailer not installed. Install with: pip install premailer")

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk, scrolledtext
    import tkinter.font as tkfont
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

# Try to import HTML rendering widgets
try:
    from tkinterweb import HtmlFrame
    TKINTERWEB_AVAILABLE = True
except ImportError:
    TKINTERWEB_AVAILABLE = False
    print("Info: tkinterweb not available for advanced preview")

try:
    import tkhtmlview
    from tkhtmlview import HTMLLabel, HTMLScrolledText
    TKHTMLVIEW_AVAILABLE = True
except ImportError:
    TKHTMLVIEW_AVAILABLE = False
    print("Info: tkhtmlview not available for HTML preview")


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('html_to_oft_converter.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class DependencyError(Exception):
    """Raised when required dependencies are missing"""
    pass


class HTMLProcessingError(Exception):
    """Raised when HTML processing fails"""
    pass


class OFTCreationError(Exception):
    """Raised when OFT file creation fails"""
    pass


class ImageHandler:
    """Handles image processing, downloading, and embedding with SSL fix"""
    
    def __init__(self, base_path: Optional[Path] = None):
        self.base_path = base_path or Path.cwd()
        self.image_cache: Dict[str, bytes] = {}
        self.session = self._create_session() if REQUESTS_AVAILABLE else None
        self.failed_images: List[str] = []
        
    def _create_session(self) -> requests.Session:
        """Create a robust requests session with retry logic and SSL disabled"""
        session = requests.Session()
        session.verify = False
        
        retry = Retry(
            total=3,
            read=3,
            connect=3,
            backoff_factor=0.3,
            status_forcelist=(500, 502, 504)
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        session.timeout = 10
        
        return session
    
    def process_image(self, src: str, for_preview: bool = False) -> Optional[Tuple[bytes, str]]:
        """
        Process image from various sources
        Args:
            src: Image source URL/path
            for_preview: If True, converts to base64 data URL for preview
        Returns: (image_data, content_type) or (base64_url, content_type) if for_preview
        """
        try:
            cache_key = hashlib.md5(src.encode()).hexdigest()
            if cache_key in self.image_cache:
                image_data = self.image_cache[cache_key]
                content_type = self._get_mime_type(src)
            else:
                # Handle data URLs
                if src.startswith('data:'):
                    image_data, content_type = self._process_data_url(src)
                # Handle remote URLs
                elif src.startswith(('http://', 'https://')):
                    image_data, content_type = self._download_image(src)
                # Handle local files
                else:
                    image_data, content_type = self._load_local_image(src)
                
                if image_data:
                    self.image_cache[cache_key] = image_data
            
            if for_preview and image_data:
                # Convert to base64 data URL for preview
                b64_data = base64.b64encode(image_data).decode('utf-8')
                data_url = f"data:{content_type};base64,{b64_data}"
                return data_url, content_type
            
            return image_data, content_type
                
        except Exception as e:
            logger.warning(f"Failed to process image {src[:100]}: {e}")
            self.failed_images.append(src)
        
        return None
    
    def _process_data_url(self, data_url: str) -> Tuple[bytes, str]:
        """Extract image data from data URL"""
        try:
            header, encoded = data_url.split(',', 1)
            content_type = header.split(':')[1].split(';')[0]
            
            if 'base64' in header:
                image_data = base64.b64decode(encoded)
            else:
                image_data = encoded.encode()
            
            return image_data, content_type
        except Exception as e:
            raise HTMLProcessingError(f"Invalid data URL: {e}")
    
    def _download_image(self, url: str) -> Tuple[bytes, str]:
        """Download image from URL with multiple fallback methods"""
        
        # Method 1: Try with requests library
        if REQUESTS_AVAILABLE and self.session:
            try:
                response = self.session.get(url, timeout=15, verify=False)
                response.raise_for_status()
                content_type = response.headers.get('Content-Type', 'image/jpeg')
                return response.content, content_type
            except Exception as e:
                logger.debug(f"Requests download failed for {url}: {e}")
        
        # Method 2: Fallback to urllib with SSL disabled
        try:
            ssl_context = ssl.create_default_context()
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE
            
            req = urllib.request.Request(
                url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            )
            
            with urllib.request.urlopen(req, context=ssl_context, timeout=15) as response:
                image_data = response.read()
                content_type = response.headers.get('Content-Type', 'image/jpeg')
                return image_data, content_type
                
        except Exception as e:
            raise HTMLProcessingError(f"Failed to download image from {url}: {e}")
    
    def _load_local_image(self, path: str) -> Tuple[bytes, str]:
        """Load image from local file system"""
        try:
            if not os.path.isabs(path):
                path = os.path.join(self.base_path, path)
            
            path = os.path.normpath(path)
            
            if not os.path.exists(path):
                raise FileNotFoundError(f"Image not found: {path}")
            
            with open(path, 'rb') as f:
                image_data = f.read()
            
            content_type = self._get_mime_type(path)
            return image_data, content_type
            
        except Exception as e:
            raise HTMLProcessingError(f"Failed to load local image {path}: {e}")
    
    @staticmethod
    def _get_mime_type(filename: str) -> str:
        """Get MIME type from filename"""
        mime_type, _ = mimetypes.guess_type(filename)
        if not mime_type:
            ext = os.path.splitext(filename)[1].lower()
            mime_map = {
                '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                '.png': 'image/png', '.gif': 'image/gif',
                '.webp': 'image/webp', '.svg': 'image/svg+xml',
                '.bmp': 'image/bmp'
            }
            mime_type = mime_map.get(ext, 'application/octet-stream')
        return mime_type


class HTMLProcessor:
    """Processes and optimizes HTML for email"""
    
    def __init__(self, html_path: Optional[Path] = None):
        self.html_path = html_path
        self.image_handler = ImageHandler(html_path.parent if html_path else None)
        self.attachments: List[Dict] = []
        self.original_html = ""
        self.processed_html = ""
        
    def process(self, html_content: str, inline_css: bool = True, 
                embed_images: bool = True, for_preview: bool = False) -> str:
        """
        Process HTML content for email
        
        Args:
            html_content: Raw HTML string
            inline_css: Whether to inline CSS styles
            embed_images: Whether to embed images as attachments
            for_preview: If True, creates preview version with base64 images
            
        Returns:
            Processed HTML string
        """
        try:
            self.original_html = html_content
            
            if not BS4_AVAILABLE:
                logger.warning("BeautifulSoup not available, returning raw HTML")
                return html_content
            
            # Parse HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Inline CSS if requested
            if inline_css and PREMAILER_AVAILABLE:
                try:
                    html_content = self._inline_css(str(soup))
                    soup = BeautifulSoup(html_content, 'html.parser')
                except Exception as e:
                    logger.warning(f"Premailer CSS inlining failed: {e}")
                    if CSSUTILS_AVAILABLE:
                        soup = self._inline_css_manual(soup)
            elif inline_css and CSSUTILS_AVAILABLE:
                soup = self._inline_css_manual(soup)
            
            # Process images
            if embed_images:
                soup = self._process_images(soup, for_preview=for_preview)
            
            # Clean up HTML
            soup = self._cleanup_html(soup)
            
            # Ensure proper encoding
            html_content = str(soup)
            self.processed_html = html_content
            
            return html_content
            
        except Exception as e:
            logger.error(f"HTML processing failed: {e}")
            raise HTMLProcessingError(f"Failed to process HTML: {e}")
    
    def get_preview_html(self) -> str:
        """Generate HTML for preview with all images as base64"""
        try:
            if not self.original_html:
                return "<html><body><p>No content to preview</p></body></html>"
            
            soup = BeautifulSoup(self.original_html, 'html.parser')
            
            # Convert all images to base64 for preview
            for img in soup.find_all('img'):
                src = img.get('src')
                if not src:
                    continue
                
                try:
                    result = self.image_handler.process_image(src, for_preview=True)
                    if result:
                        data_url, _ = result
                        img['src'] = data_url
                except Exception as e:
                    logger.debug(f"Could not convert image to base64: {src}")
            
            # Add responsive viewport meta
            if not soup.find('meta', attrs={'name': 'viewport'}):
                meta = soup.new_tag('meta')
                meta.attrs['name'] = 'viewport'
                meta.attrs['content'] = 'width=device-width, initial-scale=1.0'
                if soup.head:
                    soup.head.append(meta)
            
            # Wrap in proper HTML structure if needed
            if not soup.find('html'):
                html_tag = soup.new_tag('html')
                body_tag = soup.new_tag('body')
                body_tag.string = str(soup)
                html_tag.append(body_tag)
                soup = BeautifulSoup(str(html_tag), 'html.parser')
            
            return str(soup)
            
        except Exception as e:
            logger.error(f"Failed to generate preview HTML: {e}")
            return f"<html><body><p>Error generating preview: {e}</p></body></html>"
    
    def _inline_css(self, html: str) -> str:
        """Inline CSS using Premailer"""
        try:
            premailer = Premailer(
                html,
                strip_important=False,
                keep_style_tags=True,
                base_path=str(self.html_path.parent) if self.html_path else None,
                disable_validation=True
            )
            return premailer.transform()
        except Exception as e:
            logger.warning(f"CSS inlining failed: {e}")
            return html
    
    def _inline_css_manual(self, soup: BeautifulSoup) -> BeautifulSoup:
        """Manual CSS inlining using cssutils"""
        try:
            styles = {}
            for style_tag in soup.find_all('style'):
                css_text = style_tag.string
                if css_text:
                    sheet = cssutils.parseString(css_text)
                    for rule in sheet:
                        if rule.type == rule.STYLE_RULE:
                            selector = rule.selectorText
                            if selector not in styles:
                                styles[selector] = {}
                            for prop in rule.style:
                                styles[selector][prop.name] = prop.value
            
            for selector, style_dict in styles.items():
                try:
                    elements = soup.select(selector)
                    for element in elements:
                        current_style = element.get('style', '')
                        new_style = '; '.join([f"{k}: {v}" for k, v in style_dict.items()])
                        element['style'] = f"{current_style}; {new_style}".strip('; ')
                except Exception as e:
                    logger.debug(f"Failed to apply style for selector {selector}: {e}")
            
            return soup
            
        except Exception as e:
            logger.warning(f"Manual CSS inlining failed: {e}")
            return soup
    
    def _process_images(self, soup: BeautifulSoup, for_preview: bool = False) -> BeautifulSoup:
        """Process all images in the HTML"""
        images = soup.find_all('img')
        logger.info(f"Processing {len(images)} images...")
        
        for idx, img in enumerate(images, 1):
            src = img.get('src')
            if not src:
                continue
            
            try:
                if for_preview:
                    # For preview, convert to base64
                    result = self.image_handler.process_image(src, for_preview=True)
                    if result:
                        data_url, _ = result
                        img['src'] = data_url
                else:
                    # For OFT, process normally
                    result = self.image_handler.process_image(src)
                    if result:
                        image_data, content_type = result
                        
                        cid = f"image_{len(self.attachments)}_{hashlib.md5(src.encode()).hexdigest()[:8]}"
                        
                        self.attachments.append({
                            'data': image_data,
                            'content_type': content_type,
                            'cid': cid,
                            'filename': self._extract_filename(src)
                        })
                        
                        img['src'] = f"cid:{cid}"
                    
            except Exception as e:
                logger.warning(f"Failed to process image {src[:50]}...: {e}")
        
        if not for_preview:
            logger.info(f"Successfully embedded {len(self.attachments)} images")
        
        return soup
    
    @staticmethod
    def _extract_filename(src: str) -> str:
        """Extract filename from image source"""
        if src.startswith('data:'):
            return f"image_{hashlib.md5(src.encode()).hexdigest()[:8]}.jpg"
        
        parsed = urlparse(src)
        filename = os.path.basename(parsed.path) or 'image.jpg'
        filename = re.sub(r'[^\w\-_\. ]', '_', filename)
        
        if not os.path.splitext(filename)[1]:
            filename += '.jpg'
            
        return filename
    
    @staticmethod
    def _cleanup_html(soup: BeautifulSoup) -> BeautifulSoup:
        """Clean up HTML for email compatibility"""
        for script in soup.find_all('script'):
            script.decompose()
        
        for form in soup.find_all('form'):
            form.decompose()
        
        return soup


class OFTConverter:
    """Converts HTML to Outlook OFT template"""
    
    def __init__(self):
        if not OUTLOOK_AVAILABLE:
            raise DependencyError(
                "pywin32 is required for OFT conversion. "
                "Install with: pip install pywin32"
            )
        
        self.outlook = None
        self._initialize_outlook()
    
    def _initialize_outlook(self):
        """Initialize Outlook COM object"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            logger.info("Outlook COM object initialized successfully")
        except Exception as e:
            raise OFTCreationError(f"Failed to initialize Outlook: {e}")
    
    def create_oft(self, html_content: str, output_path: Path,
                   subject: str = "", attachments: Optional[List[Dict]] = None,
                   message_format: int = 2) -> bool:
        """Create OFT file from HTML content"""
        mail_item = None
        temp_files = []
        
        try:
            mail_item = self.outlook.CreateItem(0)
            mail_item.Subject = subject
            mail_item.BodyFormat = message_format
            mail_item.HTMLBody = html_content
            
            if attachments:
                logger.info(f"Adding {len(attachments)} attachments...")
                temp_files = self._add_attachments(mail_item, attachments)
            
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            mail_item.SaveAs(str(output_path.absolute()), 5)
            
            logger.info(f"OFT file created successfully: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to create OFT file: {e}")
            raise OFTCreationError(f"OFT creation failed: {e}")
            
        finally:
            for tmp_path in temp_files:
                try:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                except Exception as e:
                    logger.debug(f"Could not delete temp file {tmp_path}: {e}")
            
            if mail_item:
                try:
                    del mail_item
                except:
                    pass
    
    def _add_attachments(self, mail_item, attachments: List[Dict]) -> List[str]:
        """Add attachments to mail item and return temp file paths"""
        temp_files = []
        
        for idx, att_info in enumerate(attachments, 1):
            try:
                filename = att_info['filename']
                ext = os.path.splitext(filename)[1] or '.jpg'
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                    tmp_file.write(att_info['data'])
                    tmp_path = tmp_file.name
                    temp_files.append(tmp_path)
                
                attachment = mail_item.Attachments.Add(tmp_path)
                
                if 'cid' in att_info:
                    pr_attach_content_id = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                    attachment.PropertyAccessor.SetProperty(
                        pr_attach_content_id,
                        att_info['cid']
                    )
                
                logger.debug(f"Added attachment {idx}/{len(attachments)}: {filename}")
                    
            except Exception as e:
                logger.warning(f"Failed to add attachment {att_info.get('filename')}: {e}")
        
        return temp_files


class HTMLToOFTConverter:
    """Main converter class orchestrating the conversion process"""
    
    def __init__(self):
        self.html_processor = None
        self.oft_converter = None
        self._check_dependencies()
    
    def _check_dependencies(self):
        """Check for required dependencies"""
        missing = []
        
        if not OUTLOOK_AVAILABLE:
            missing.append("pywin32")
        if not BS4_AVAILABLE:
            missing.append("beautifulsoup4")
        
        if missing:
            raise DependencyError(
                f"Missing required dependencies: {', '.join(missing)}\n"
                f"Install with: pip install {' '.join(missing)}"
            )
    
    def convert_file(self, input_path: str, output_path: Optional[str] = None,
                     subject: str = "", inline_css: bool = True,
                     embed_images: bool = True) -> Path:
        """Convert HTML file to OFT"""
        try:
            input_path = Path(input_path)
            
            if not input_path.exists():
                raise FileNotFoundError(f"Input file not found: {input_path}")
            
            if not input_path.suffix.lower() in ['.html', '.htm']:
                raise ValueError("Input file must be HTML (.html or .htm)")
            
            if output_path is None:
                output_path = input_path.with_suffix('.oft')
            else:
                output_path = Path(output_path)
            
            logger.info(f"Converting {input_path} to {output_path}")
            
            # Read HTML content
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
            html_content = None
            
            for encoding in encodings:
                try:
                    with open(input_path, 'r', encoding=encoding) as f:
                        html_content = f.read()
                    logger.debug(f"Successfully read file with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            
            if html_content is None:
                raise ValueError("Could not read HTML file with any supported encoding")
            
            # Process HTML
            self.html_processor = HTMLProcessor(input_path)
            processed_html = self.html_processor.process(
                html_content,
                inline_css=inline_css,
                embed_images=embed_images,
                for_preview=False
            )
            
            # Create OFT
            self.oft_converter = OFTConverter()
            self.oft_converter.create_oft(
                processed_html,
                output_path,
                subject=subject,
                attachments=self.html_processor.attachments
            )
            
            logger.info(f"✓ Conversion completed successfully: {output_path}")
            
            # Print summary
            total_images = len(self.html_processor.image_handler.failed_images) + len(self.html_processor.attachments)
            if total_images > 0:
                logger.info(f"Image Summary: {len(self.html_processor.attachments)}/{total_images} images embedded successfully")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Conversion failed: {e}")
            raise
    
    def get_preview_html(self) -> str:
        """Get preview HTML from last conversion"""
        if self.html_processor:
            return self.html_processor.get_preview_html()
        return "<html><body><p>No conversion performed yet</p></body></html>"


class EmailPreviewWindow:
    """Dedicated window for previewing email templates"""
    
    def __init__(self, parent, html_content: str, title: str = "Email Template Preview"):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("900x700")
        
        self.html_content = html_content
        self.temp_file = None
        
        self._setup_ui()
        self._load_preview()
    
    def _setup_ui(self):
        """Setup the preview window UI"""
        # Toolbar
        toolbar = ttk.Frame(self.window)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Label(toolbar, text="Preview Mode:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            toolbar,
            text="🔄 Refresh",
            command=self._load_preview,
            width=12
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="🌐 Open in Browser",
            command=self._open_in_browser,
            width=18
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="💾 Save HTML",
            command=self._save_html,
            width=15
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="📋 Copy HTML",
            command=self._copy_html,
            width=15
        ).pack(side=tk.LEFT, padx=2)
        
        # Separator
        ttk.Separator(self.window, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5)
        
        # Preview area
        preview_frame = ttk.Frame(self.window)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Try different preview methods in order of preference
        self.preview_widget = None
        
        # Method 1: Try tkinterweb (best - full browser engine)
        if TKINTERWEB_AVAILABLE:
            try:
                self.preview_widget = HtmlFrame(preview_frame, messages_enabled=False)
                self.preview_method = "tkinterweb"
                logger.info("Using tkinterweb for preview (full browser rendering)")
            except Exception as e:
                logger.debug(f"tkinterweb initialization failed: {e}")
        
        # Method 2: Try tkhtmlview (good - HTML subset)
        if not self.preview_widget and TKHTMLVIEW_AVAILABLE:
            try:
                self.preview_widget = HTMLScrolledText(
                    preview_frame,
                    html="<html><body>Loading...</body></html>",
                    wrap=tk.WORD
                )
                self.preview_method = "tkhtmlview"
                logger.info("Using tkhtmlview for preview (HTML rendering)")
            except Exception as e:
                logger.debug(f"tkhtmlview initialization failed: {e}")
        
        # Method 3: Fallback to scrolled text (basic - no rendering)
        if not self.preview_widget:
            self.preview_widget = scrolledtext.ScrolledText(
                preview_frame,
                wrap=tk.WORD,
                font=('Courier', 9)
            )
            self.preview_method = "text"
            logger.info("Using plain text for preview (HTML source)")
        
        self.preview_widget.pack(fill=tk.BOTH, expand=True)
        
        # Status bar
        status_frame = ttk.Frame(self.window)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        
        self.status_label = ttk.Label(
            status_frame,
            text=f"Preview Method: {self.preview_method.upper()} | Size: {len(self.html_content):,} bytes",
            font=('Arial', 8)
        )
        self.status_label.pack(side=tk.LEFT)
        
        ttk.Button(
            status_frame,
            text="Close",
            command=self.window.destroy,
            width=10
        ).pack(side=tk.RIGHT, padx=5)
    
    def _load_preview(self):
        """Load the HTML preview"""
        try:
            if self.preview_method == "tkinterweb":
                # Load HTML directly into tkinterweb
                self.preview_widget.load_html(self.html_content)
                
            elif self.preview_method == "tkhtmlview":
                # Load HTML into tkhtmlview
                self.preview_widget.set_html(self.html_content)
                
            else:
                # Show HTML source in text widget
                self.preview_widget.delete(1.0, tk.END)
                self.preview_widget.insert(1.0, self.html_content)
                
                # Basic syntax highlighting
                self._highlight_html()
            
            logger.info("Preview loaded successfully")
            
        except Exception as e:
            logger.error(f"Failed to load preview: {e}")
            messagebox.showerror("Preview Error", f"Failed to load preview:\n{e}")
    
    def _highlight_html(self):
        """Basic syntax highlighting for HTML in text widget"""
        if self.preview_method != "text":
            return
        
        try:
            # Tag configuration
            self.preview_widget.tag_config("tag", foreground="#0066CC")
            self.preview_widget.tag_config("attribute", foreground="#CC6600")
            self.preview_widget.tag_config("string", foreground="#009900")
            
            # Find and highlight tags
            import re
            
            # HTML tags
            for match in re.finditer(r'<[^>]+>', self.html_content):
                start_idx = f"1.0+{match.start()}c"
                end_idx = f"1.0+{match.end()}c"
                self.preview_widget.tag_add("tag", start_idx, end_idx)
            
        except Exception as e:
            logger.debug(f"Syntax highlighting failed: {e}")
    
    def _open_in_browser(self):
        """Open preview in default web browser"""
        try:
            # Create temporary HTML file
            if not self.temp_file:
                with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                    f.write(self.html_content)
                    self.temp_file = f.name
            
            # Open in browser
            webbrowser.open('file://' + os.path.abspath(self.temp_file))
            logger.info(f"Opened preview in browser: {self.temp_file}")
            
        except Exception as e:
            logger.error(f"Failed to open in browser: {e}")
            messagebox.showerror("Error", f"Failed to open in browser:\n{e}")
    
    def _save_html(self):
        """Save HTML to file"""
        try:
            filename = filedialog.asksaveasfilename(
                title="Save HTML Preview",
                defaultextension=".html",
                filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
            )
            
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.html_content)
                
                logger.info(f"HTML saved to: {filename}")
                messagebox.showinfo("Success", f"HTML saved to:\n{filename}")
                
        except Exception as e:
            logger.error(f"Failed to save HTML: {e}")
            messagebox.showerror("Error", f"Failed to save HTML:\n{e}")
    
    def _copy_html(self):
        """Copy HTML to clipboard"""
        try:
            self.window.clipboard_clear()
            self.window.clipboard_append(self.html_content)
            self.window.update()
            
            self.status_label.config(text="✓ HTML copied to clipboard")
            logger.info("HTML copied to clipboard")
            
        except Exception as e:
            logger.error(f"Failed to copy HTML: {e}")
            messagebox.showerror("Error", f"Failed to copy HTML:\n{e}")
    
    def __del__(self):
        """Cleanup temporary files"""
        if self.temp_file and os.path.exists(self.temp_file):
            try:
                os.unlink(self.temp_file)
            except:
                pass


class ConverterGUI:
    """Enhanced GUI with preview functionality"""
    
    def __init__(self):
        if not TKINTER_AVAILABLE:
            raise DependencyError("tkinter is required for GUI")
        
        self.converter = HTMLToOFTConverter()
        self.root = tk.Tk()
        self.root.title("HTML to OFT Converter - Pro Edition v3.0 with Preview")
        self.root.geometry("1000x750")
        
        # Preview window reference
        self.preview_window = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the user interface"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        title_label = ttk.Label(
            header_frame,
            text="HTML to Outlook OFT Converter",
            font=('Arial', 18, 'bold'),
            foreground='#0066CC'
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Professional Email Template Converter with Live Preview",
            font=('Arial', 10),
            foreground='#666666'
        )
        subtitle_label.pack()
        
        # Version badge
        version_label = ttk.Label(
            header_frame,
            text="v3.0 Pro",
            font=('Arial', 8, 'bold'),
            foreground='white',
            background='#0066CC',
            padding=(5, 2)
        )
        version_label.pack(pady=5)
        
        # Input section
        row = 1
        input_frame = ttk.LabelFrame(main_frame, text="📁 Input / Output Files", padding="10")
        input_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(1, weight=1)
        
        # Input file
        ttk.Label(input_frame, text="HTML File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_var, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self._browse_input).grid(row=0, column=2, padx=5, pady=5)
        
        # Output file
        ttk.Label(input_frame, text="OFT Output:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.output_var, width=60).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(input_frame, text="Browse...", command=self._browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Subject
        ttk.Label(input_frame, text="Email Subject:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.subject_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.subject_var, width=60).grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        # Options section
        row += 1
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ Processing Options", padding="10")
        options_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        options_left = ttk.Frame(options_frame)
        options_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.inline_css_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_left,
            text="✓ Inline CSS styles (recommended for email clients)",
            variable=self.inline_css_var
        ).pack(anchor=tk.W, pady=2)
        
        self.embed_images_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_left,
            text="✓ Embed images (download and attach remote images)",
            variable=self.embed_images_var
        ).pack(anchor=tk.W, pady=2)
        
        # Info
        info_label = ttk.Label(
            options_frame,
            text="ℹ️ SSL verification disabled for corporate firewalls",
            font=('Arial', 8),
            foreground='#666666'
        )
        info_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Action buttons
        row += 1
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=15)
        
        self.convert_btn = ttk.Button(
            button_frame,
            text="🔄 Convert to OFT",
            command=self._convert_single,
            width=20
        )
        self.convert_btn.grid(row=0, column=0, padx=5)
        
        self.preview_btn = ttk.Button(
            button_frame,
            text="👁️ Preview Template",
            command=self._show_preview,
            width=20,
            state='disabled'
        )
        self.preview_btn.grid(row=0, column=1, padx=5)
        
        self.batch_btn = ttk.Button(
            button_frame,
            text="📂 Batch Convert",
            command=self._batch_convert,
            width=18
        )
        self.batch_btn.grid(row=0, column=2, padx=5)
        
        ttk.Button(
            button_frame,
            text="🗑️ Clear",
            command=self._clear_form,
            width=12
        ).grid(row=0, column=3, padx=5)
        
        # Quick actions
        quick_frame = ttk.Frame(button_frame)
        quick_frame.grid(row=1, column=0, columnspan=4, pady=(10, 0))
        
        ttk.Button(
            quick_frame,
            text="📖 Help",
            command=self._show_help,
            width=12
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            quick_frame,
            text="ℹ️ About",
            command=self._show_about,
            width=12
        ).pack(side=tk.LEFT, padx=2)
        
        # Progress section
        row += 1
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="Ready to convert")
        status_label = ttk.Label(
            progress_frame,
            textvariable=self.status_var,
            font=('Arial', 10, 'bold'),
            foreground='#0066CC'
        )
        status_label.pack(pady=5)
        
        # Log section
        row += 1
        log_frame = ttk.LabelFrame(main_frame, text="📋 Conversion Log", padding="5")
        log_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(row, weight=1)
        
        # Log text with scrollbar
        log_container = ttk.Frame(log_frame)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_container,
            height=12,
            width=90,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Log controls
        log_controls = ttk.Frame(log_frame)
        log_controls.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(
            log_controls,
            text="Clear Log",
            command=lambda: self.log_text.delete(1.0, tk.END)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            log_controls,
            text="Save Log",
            command=self._save_log
        ).pack(side=tk.LEFT, padx=2)
        
        # Redirect logger
        self._setup_logging()
    
    def _setup_logging(self):
        """Setup logging to GUI"""
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
            
            def emit(self, record):
                msg = self.format(record)
                
                def append():
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.see(tk.END)
                    
                    # Color code by level
                    if record.levelname == 'ERROR':
                        # Color last line red
                        pass
                    elif record.levelname == 'WARNING':
                        # Color last line orange
                        pass
                
                self.text_widget.after(0, append)
        
        handler = TextHandler(self.log_text)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S'))
        logger.addHandler(handler)
    
    def _browse_input(self):
        """Browse for input file"""
        filename = filedialog.askopenfilename(
            title="Select HTML Email Template",
            filetypes=[
                ("HTML files", "*.html *.htm"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.input_var.set(filename)
            
            if not self.output_var.get():
                output = Path(filename).with_suffix('.oft')
                self.output_var.set(str(output))
            
            # Enable preview button
            self.preview_btn.config(state='normal')
    
    def _browse_output(self):
        """Browse for output file"""
        filename = filedialog.asksaveasfilename(
            title="Save OFT Template As",
            defaultextension=".oft",
            filetypes=[
                ("Outlook Template", "*.oft"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.output_var.set(filename)
    
    def _convert_single(self):
        """Convert single file"""
        input_file = self.input_var.get()
        output_file = self.output_var.get()
        subject = self.subject_var.get()
        
        if not input_file:
            messagebox.showerror("Missing Input", "Please select an input HTML file")
            return
        
        if not output_file:
            messagebox.showerror("Missing Output", "Please specify an output OFT file location")
            return
        
        # Disable buttons
        self._set_buttons_state('disabled')
        
        try:
            self.progress.start()
            self.status_var.set("🔄 Converting... Please wait")
            self.root.update()
            
            result_path = self.converter.convert_file(
                input_file,
                output_file,
                subject=subject,
                inline_css=self.inline_css_var.get(),
                embed_images=self.embed_images_var.get()
            )
            
            self.progress.stop()
            self.status_var.set("✅ Conversion completed successfully!")
            
            # Enable preview
            self.preview_btn.config(state='normal')
            
            # Show success message with options
            response = messagebox.askyesno(
                "Success!",
                f"OFT file created successfully!\n\n{result_path}\n\n"
                f"Would you like to preview the template?",
                icon='info'
            )
            
            if response:
                self._show_preview()
            
        except Exception as e:
            self.progress.stop()
            self.status_var.set("❌ Conversion failed")
            messagebox.showerror("Conversion Error", f"Failed to convert:\n\n{str(e)}")
        
        finally:
            self._set_buttons_state('normal')
    
    def _show_preview(self):
        """Show template preview"""
        input_file = self.input_var.get()
        
        if not input_file:
            messagebox.showwarning("No File", "Please select an HTML file first")
            return
        
        try:
            self.status_var.set("🔄 Generating preview...")
            self.root.update()
            
            # Read HTML file
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
            html_content = None
            
            for encoding in encodings:
                try:
                    with open(input_file, 'r', encoding=encoding) as f:
                        html_content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            
            if not html_content:
                raise ValueError("Could not read HTML file")
            
            # Process for preview
            processor = HTMLProcessor(Path(input_file))
            preview_html = processor.process(
                html_content,
                inline_css=self.inline_css_var.get(),
                embed_images=self.embed_images_var.get(),
                for_preview=True
            )
            
            # Close existing preview window if open
            if self.preview_window:
                try:
                    self.preview_window.window.destroy()
                except:
                    pass
            
            # Open new preview window
            self.preview_window = EmailPreviewWindow(
                self.root,
                preview_html,
                f"Preview: {Path(input_file).name}"
            )
            
            self.status_var.set("✅ Preview opened")
            
        except Exception as e:
            self.status_var.set("❌ Preview failed")
            messagebox.showerror("Preview Error", f"Failed to generate preview:\n\n{str(e)}")
            logger.error(f"Preview generation failed: {e}")
    
    def _batch_convert(self):
        """Batch convert multiple files"""
        input_dir = filedialog.askdirectory(title="Select folder containing HTML files")
        if not input_dir:
            return
        
        output_dir = filedialog.askdirectory(title="Select output folder for OFT files")
        if not output_dir:
            return
        
        self._set_buttons_state('disabled')
        
        try:
            self.progress.start()
            self.status_var.set("🔄 Batch converting... Please wait")
            self.root.update()
            
            results = self.converter.batch_convert(
                input_dir,
                output_dir,
                inline_css=self.inline_css_var.get(),
                embed_images=self.embed_images_var.get()
            )
            
            self.progress.stop()
            
            success_count = sum(1 for _, success, _ in results if success)
            total_count = len(results)
            
            self.status_var.set(f"✅ Batch complete: {success_count}/{total_count} successful")
            
            # Show detailed summary
            summary = f"Batch Conversion Results:\n\n"
            summary += f"Total files: {total_count}\n"
            summary += f"Successful: {success_count}\n"
            summary += f"Failed: {total_count - success_count}\n\n"
            
            failed = [(path, error) for path, success, error in results if not success]
            if failed:
                summary += "Failed files:\n"
                for path, error in failed[:5]:
                    summary += f"  • {Path(path).name}\n"
                if len(failed) > 5:
                    summary += f"  ... and {len(failed) - 5} more\n"
            
            messagebox.showinfo("Batch Conversion Complete", summary)
            
        except Exception as e:
            self.progress.stop()
            self.status_var.set("❌ Batch conversion failed")
            messagebox.showerror("Error", f"Batch conversion failed:\n\n{str(e)}")
        
        finally:
            self._set_buttons_state('normal')
    
    def _clear_form(self):
        """Clear the form"""
        self.input_var.set("")
        self.output_var.set("")
        self.subject_var.set("")
        self.status_var.set("Ready to convert")
        self.preview_btn.config(state='disabled')
    
    def _save_log(self):
        """Save log to file"""
        try:
            filename = filedialog.asksaveasfilename(
                title="Save Log File",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if filename:
                log_content = self.log_text.get(1.0, tk.END)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                
                messagebox.showinfo("Success", f"Log saved to:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save log:\n{e}")
    
    def _show_help(self):
        """Show help dialog"""
        help_text = """
HTML to OFT Converter - Quick Guide

1. SELECT INPUT FILE
   • Click 'Browse' next to HTML File
   • Choose your email template (.html or .htm)

2. CHOOSE OUTPUT LOCATION
   • Click 'Browse' next to OFT Output
   • Select where to save the .oft file

3. CONFIGURE OPTIONS
   • Inline CSS: Recommended for email compatibility
   • Embed Images: Downloads and attaches remote images

4. CONVERT
   • Click 'Convert to OFT' to create the template
   • Use 'Preview Template' to see the result

5. PREVIEW FEATURES
   • View your email template before sending
   • Open in browser for full preview
   • Save or copy HTML code

BATCH CONVERSION
   • Use 'Batch Convert' for multiple files
   • Select input folder with HTML files
   • Choose output folder for OFT files

TIPS
   • SSL verification is disabled for corporate networks
   • Images are automatically embedded
   • Templates work in all Outlook versions
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Help - HTML to OFT Converter")
        help_window.geometry("600x500")
        
        text_widget = scrolledtext.ScrolledText(
            help_window,
            wrap=tk.WORD,
            font=('Courier', 10),
            padx=10,
            pady=10
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(1.0, help_text)
        text_widget.config(state='disabled')
        
        ttk.Button(
            help_window,
            text="Close",
            command=help_window.destroy
        ).pack(pady=10)
    
    def _show_about(self):
        """Show about dialog"""
        about_text = f"""
HTML to OFT Converter
Professional Edition v3.0

Features:
• Convert HTML email templates to Outlook OFT format
• Live template preview (no Outlook Premium required!)
• Automatic image embedding
• CSS inlining for email compatibility
• Batch conversion support
• SSL/TLS handling for corporate networks

Preview Methods:
• tkinterweb: Full browser rendering ({'Available' if TKINTERWEB_AVAILABLE else 'Not installed'})
• tkhtmlview: HTML subset rendering ({'Available' if TKHTMLVIEW_AVAILABLE else 'Not installed'})
• Plain text: HTML source view (Always available)

Dependencies Status:
• pywin32: {'✓ Installed' if OUTLOOK_AVAILABLE else '✗ Not installed'}
• BeautifulSoup4: {'✓ Installed' if BS4_AVAILABLE else '✗ Not installed'}
• requests: {'✓ Installed' if REQUESTS_AVAILABLE else '✗ Not installed'}
• premailer: {'✓ Installed' if PREMAILER_AVAILABLE else '✗ Not installed'}

For support or issues, check the log file:
html_to_oft_converter.log

© 2024 Advanced Email Engineering Team
        """
        
        messagebox.showinfo("About", about_text)
    
    def _set_buttons_state(self, state):
        """Enable or disable all buttons"""
        self.convert_btn.config(state=state)
        self.batch_btn.config(state=state)
        if state == 'normal' and self.input_var.get():
            self.preview_btn.config(state=state)
    
    def run(self):
        """Run the GUI"""
        self.root.mainloop()


class CLI:
    """Command-line interface"""
    
    @staticmethod
    def run():
        import argparse
        
        parser = argparse.ArgumentParser(
            description='HTML to OFT Converter v3.0 - With Preview Support',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
  # Launch GUI
  python html_to_oft_v3.py --gui
  
  # Convert single file
  python html_to_oft_v3.py input.html -o output.oft -s "Subject"
  
  # Batch convert
  python html_to_oft_v3.py -d input_folder -o output_folder
  
  # Generate preview HTML
  python html_to_oft_v3.py input.html --preview-only
            """
        )
        
        parser.add_argument('input', nargs='?', help='Input HTML file or directory')
        parser.add_argument('-o', '--output', help='Output OFT file or directory')
        parser.add_argument('-s', '--subject', default='', help='Email subject')
        parser.add_argument('-d', '--directory', action='store_true', help='Batch process directory')
        parser.add_argument('--no-inline-css', action='store_true', help='Do not inline CSS')
        parser.add_argument('--no-embed-images', action='store_true', help='Do not embed images')
        parser.add_argument('--preview-only', action='store_true', help='Generate preview HTML only')
        parser.add_argument('--gui', action='store_true', help='Launch GUI')
        parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')
        
        args = parser.parse_args()
        
        if args.verbose:
            logger.setLevel(logging.DEBUG)
        
        if args.gui:
            try:
                gui = ConverterGUI()
                gui.run()
            except Exception as e:
                print(f"Failed to launch GUI: {e}")
                sys.exit(1)
            return
        
        if not args.input:
            parser.print_help()
            sys.exit(1)
        
        try:
            converter = HTMLToOFTConverter()
            
            if args.preview_only:
                # Generate preview only
                with open(args.input, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                processor = HTMLProcessor(Path(args.input))
                preview_html = processor.process(html_content, for_preview=True)
                
                output_file = args.output or str(Path(args.input).with_suffix('.preview.html'))
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(preview_html)
                
                print(f"✓ Preview HTML generated: {output_file}")
                
            elif args.directory:
                results = converter.batch_convert(
                    args.input,
                    args.output,
                    inline_css=not args.no_inline_css,
                    embed_images=not args.no_embed_images
                )
                
                success_count = sum(1 for _, success, _ in results if success)
                print(f"\n{'='*60}")
                print(f"Batch Conversion Summary")
                print(f"{'='*60}")
                print(f"Total: {len(results)} | Success: {success_count} | Failed: {len(results) - success_count}")
                
            else:
                output_path = converter.convert_file(
                    args.input,
                    args.output,
                    subject=args.subject,
                    inline_css=not args.no_inline_css,
                    embed_images=not args.no_embed_images
                )
                
                print(f"\n{'='*60}")
                print(f"✓ Conversion Successful")
                print(f"{'='*60}")
                print(f"Output: {output_path}")
        
        except Exception as e:
            print(f"\n{'='*60}")
            print(f"✗ Error: {e}")
            print(f"{'='*60}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            sys.exit(1)


def main():
    """Main entry point"""
    print("="*70)
    print("HTML to OFT Converter v3.0 - Professional Edition with Preview")
    print("="*70)
    print()
    CLI.run()


if __name__ == '__main__':

    main()

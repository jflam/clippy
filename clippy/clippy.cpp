// Console mode program that writes any bitmap contents out to disk using 
// either a png or a jpg encoder.

#include "pch.h"
#include <string>
#include <iostream>
#include <iomanip>

// Helper method to convert a UTF-8 encoded string into a STL wstring
std::wstring s2ws(const std::string& s)
{
	int len;
	int slength = (int)s.length() + 1;
	len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0);
	wchar_t* buf = new wchar_t[len];
	MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, buf, len);
	std::wstring r(buf);
	delete[] buf;
	return r;
}

// Helper method that generates an error message + 32 bit hex representation of the HRESULT
void RaiseError(std::string errorMessage, HRESULT hr)
{
	std::cout << errorMessage << ": HRESULT = 0x" << std::uppercase << std::setfill('0') << std::setw(8) << std::hex << hr;
}

HRESULT WriteBitmapToDisk(std::string filename, 
						  GUID encoderId,
						  int output_width,
						  int output_height,
						  IWICImagingFactory* ipFactory, 
						  IWICBitmapSource* ipBitmapSource)
{
	HRESULT hr;
	IWICStream *ipStream = NULL;
	IWICBitmapEncoder *ipBitmapEncoder = NULL;
	IWICBitmapFrameEncode *ipFrameEncoder = NULL;

	// Create the appropriate WIC Bitmap Encoder object based on whether the user wants the file to be 
	// serialized as a PNG or a JPEG. TODO: in the future perhaps have an "auto" mode that attempts to serialize
	// to an in-memory buffer to see which mechanism results in smaller images? 
	// TODO: perhaps write a macro that randomly injects failing HRESULTs into debug builds to test the
	// recovery code paths?
	hr = ipFactory->CreateEncoder(encoderId,
		NULL,
		&ipBitmapEncoder);
	if (FAILED(hr))
	{
		RaiseError("Could not create the PNG or JPG encoder: ", hr);
		goto FreeCOM;
	}

	// Construct a WIC stream object using the factory
	hr = ipFactory->CreateStream(&ipStream);
	if (FAILED(hr)) {
		RaiseError("Could not create an IStream object: ", hr);
		goto FreeCOM;
	}

	// Initialize the WIC stream to write its contents out to the filename specified below. Ensure that
	// the correct filename extension (png|jpg) is appended
	hr = ipStream->InitializeFromFilename(s2ws(filename).c_str(), GENERIC_WRITE);
	if (FAILED(hr))
	{
		RaiseError("Failed to initialize a writeable stream: ", hr);
		goto FreeCOM;
	}

	// Tell the WIC Bitmap Encoder to write to the WIC stream object
	hr = ipBitmapEncoder->Initialize(ipStream, WICBitmapEncoderNoCache);
	if (FAILED(hr))
	{
		RaiseError("Failed to initialize the bitmap encoder using the stream: ", hr);
		goto FreeCOM;
	}

	// Construct a WIC Frame Encoder object that we will be using to encode the WIC Bitmap object
	hr = ipBitmapEncoder->CreateNewFrame(&ipFrameEncoder, NULL);
	if (FAILED(hr))
	{
		RaiseError("Failed to create a new frame encoder using the bitmap encoder: ", hr);
		goto FreeCOM;
	}

	// Initialize the WIC Frame Encoder object. This is required otherwise subsequent calls will fail
	hr = ipFrameEncoder->Initialize(NULL);
	if (FAILED(hr))
	{
		RaiseError("Failed to initialize the frame encoder: ", hr);
		goto FreeCOM;
	}

	// Set the size of the frame encoder to be the final size of the image
	hr = ipFrameEncoder->SetSize(output_width, output_height);
	if (FAILED(hr))
	{
		RaiseError("Failed to set the output size for the frame encoder: ", hr);
		goto FreeCOM;
	}

	// Set the correct pixel format. This is RGB 8 bits per color channel.
	WICPixelFormatGUID formatGuid;
	formatGuid = GUID_WICPixelFormat32bppRGBA; //GUID_WICPixelFormat24bppRGB;
	hr = ipFrameEncoder->SetPixelFormat(&formatGuid);
	if (FAILED(hr))
	{
		RaiseError("Failed to set the pixel format (WICPixelFormat24bppRGB) for the frame encoder: ", hr);
		goto FreeCOM;
	}

	// Tell the WIC Frame Encoder to use the WIC Bitmap Scaler object we created earlier.
	// This completes the construction of the pipeline:
	// Bitmap -> BitmapScaler -> PngEncoder -> FrameEncoder -> Stream
	hr = ipFrameEncoder->WriteSource(ipBitmapSource, NULL);
	if (FAILED(hr))
	{
		RaiseError("Failed to set the write source of the frame encoder to the bitmap scaler: ", hr);
		goto FreeCOM;
	}

	// Tell the WIC Frame Encoder to serialize the frame to the stream
	hr = ipFrameEncoder->Commit();
	if (FAILED(hr))
	{
		RaiseError("Failed to commit the frame encoder: ", hr);
		goto FreeCOM;
	}

	// Tell the WIC Bitmap Encoder to serialize the image (which includes a frame) to the stream
	hr = ipBitmapEncoder->Commit();
	if (FAILED(hr))
	{
		RaiseError("Failed to commit the bitmap encoder: ", hr);
		goto FreeCOM;
	}

FreeCOM:
	if (ipStream != NULL)
	{
		ipStream->Release();
	}

	if (ipFrameEncoder != NULL)
	{
		ipFrameEncoder->Release();
	}

	if (ipBitmapEncoder != NULL)
	{
		ipBitmapEncoder->Release();
	}

	return hr;
}

int main(int argc, char* argv[])
{
	cxxopts::Options options("clippy", "Write clipboard bitmap to disk as a file");
	options.add_options()
		("max_width", "Maximum width of image (defaults to 800)", cxxopts::value<int>()->default_value("800"))
		("write_full", "Write full sized image in addition to resized image to disk", cxxopts::value<bool>()->default_value("false"))
		("f,filename", "Filename to write, but without extension (defaults to 'image')", cxxopts::value<std::string>()->default_value("image"))
		("encoder", "Bitmap encoder to use: (png|jpeg, defaults to png)", cxxopts::value<std::string>()->default_value("png"))
		("test_clipboard_has_bitmap", "If true, only tests to see if clipboard contains a bitmap. Writes TRUE to stdout if it does", cxxopts::value<bool>()->default_value("false"));

	// TODO: validate the parameters
	auto result = options.parse(argc, argv);
	std::string filename = result["filename"].as<std::string>();
	int max_width = result["max_width"].as<int>();
	bool write_full = result["write_full"].as<bool>();
	std::string encoder = result["encoder"].as<std::string>();
	bool test_clipboard_has_bitmap = result["test_clipboard_has_bitmap"].as<bool>();

	GUID encoderId = encoder == "png" ? GUID_ContainerFormatPng : GUID_ContainerFormatJpeg;

	HBITMAP hBitmap;
	HRESULT hr;

	IWICImagingFactory *ipFactory = NULL;
	IWICBitmapScaler *ipScaler = NULL;
	IWICBitmap *ipBitmap = NULL;

	// Attempt to open the Clipboard object. Failure will result in exit. A successful call must be balanced by
	// a call to CloseClipboard()
	if (!OpenClipboard(NULL))
	{
		RaiseError("Failed to open the clipboard object", E_FAIL);
		return 1;
	}

	// Attempt to retrieve a HBITMAP object from the clipboard. Failure is OK - there is no bitmap to write to disk.
	hBitmap = reinterpret_cast<HBITMAP>(GetClipboardData(CF_BITMAP));

	// If caller is looking for the presence of a bitmap on the clipboard, return TRUE or FALSE and exit normally
	if (test_clipboard_has_bitmap)
	{
		std::cout << (hBitmap == NULL ? "FALSE" : "TRUE");
		return 0;
	}

	// Caller wants to write bitmap to disk, so a missing bitmap is a failure exit
	if (hBitmap == NULL)
	{
		RaiseError("No bitmap on clipboard", E_FAIL);
		goto FreeClipboard;
	}

	// Must initialize COM on this thread before we can use it. A successful call must be balanced by a call to
	// CoUninitialize() which will unload inproc servers from the process as part of cleanup. TODO: We may want to
	// have an explicit init / uninit protocol (perhaps with the creation of the Clipboard object) in the future
	// so we don't thrash the OS loading and unloading the WIC module.
	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		RaiseError("Failed to initialize COM: ", hr);
		goto FreeClipboard;
	}

	// The WIC Imaging Factory object is used to create resources that we will be using:
	// 1. A WIC Bitmap object that we create from the HBITMAP object that we received from the clipboard
	// 2. A WIC Bitmap Scaler object that we will use to resize the bitmap to the size determined by width_constraint
	// 3. A WIC Stream object that is used to write the serialized bitmap(s) to disk
	// 4. A WIC Bitmap Encoder (either for PNG or JPG) that is used to serialize the bitmap(s) to disk
	hr = CoCreateInstance(CLSID_WICImagingFactory,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IWICImagingFactory,
		reinterpret_cast<void **>(&ipFactory));
	if (FAILED(hr))
	{
		RaiseError("Failed to initialize WIC Imaging Factory object: ", hr);

		// Notice that COM is initialized here, so we go to the section where we release all COM interfaces
		// that have been explicitly assigned
		goto FreeCOM;
	}

	// Bitmap objects stored on the clipboard are referenced by a generic Windows HANDLE object. Here we 
	// initialize the WIC Bitmap object using the data stored in the HBITMAP.
	hr = ipFactory->CreateBitmapFromHBITMAP(reinterpret_cast<HBITMAP>(hBitmap),
		NULL,
		WICBitmapIgnoreAlpha,
		&ipBitmap);
	if (FAILED(hr))
	{
		RaiseError("Failed to construct a WIC Bitmap object from the HBITMAP from the clipboard: ", hr);
		goto FreeCOM;
	}

	// Retrieve the width and height of the clipboard image
	UINT width, height;
	hr = ipBitmap->GetSize(&width, &height);
	if (FAILED(hr))
	{
		RaiseError("Could not get the width and height of the WIC Bitmap object: ", hr);
		goto FreeCOM;
	}

	// Write full-sized image
	if (write_full)
	{
		hr = WriteBitmapToDisk(filename + "_full." + encoder, encoderId, width, height, ipFactory, ipBitmap);
		if (FAILED(hr))
		{
			RaiseError("Could not write full sized image to disk: ", hr);
			goto FreeCOM;
		}
	}

	// Write resized image

	// Compute the output image width and height by constraining the maximum width of the image to max_width
	UINT output_width, output_height;
	float scaling_factor;
	scaling_factor = (float)((float)max_width / (float)width);
	output_width = max_width;
	output_height = (UINT)(scaling_factor * height);

	// Now resize it using a WIC Bitmap Scaler object
	hr = ipFactory->CreateBitmapScaler(&ipScaler);
	if (FAILED(hr))
	{
		RaiseError("Could not create a WIC Bitmap scaler object: ", hr);
		goto FreeCOM;
	}

	// Create a WIC Bitmap Scaler object that we will use to resize the object. In this case note that the
	// HighQualityCubic interpolation option only ships in Windows 10. TODO: perhaps have a fallback option
	// here in case the user is running on an older OS (we really need to test on an older OS to see if this is true)
	hr = ipScaler->Initialize(ipBitmap,
		output_width,
		output_height,
		WICBitmapInterpolationModeHighQualityCubic);
	if (FAILED(hr))
	{
		RaiseError("Could not initialize the WIC Bitmap scaler object InterpolationMode High Quality Cubic: ", hr);
		goto FreeCOM;
	}

	hr = WriteBitmapToDisk(filename + "." + encoder, encoderId, output_width, output_height, ipFactory, ipScaler);
	if (FAILED(hr))
	{
		RaiseError("Could not write resized image to disk: ", hr);
		goto FreeCOM;
	}

FreeCOM:
	// Free all COM interfaces that have been assigned. There should only be a single AddRef to any
	// of these interfaces, so the single Release if unassigned will do the right thing.
	if (ipScaler != NULL)
	{
		ipScaler->Release();
	}

	if (ipFactory != NULL)
	{
		ipFactory->Release();
	}

	// Turn off COM for this thread - this will likely unload the WIC component, so TODO we really 
	// should move this to the destructor of the Clipboard class when we write that.
	CoUninitialize();

FreeClipboard:
	// Ensure that we have closed the clipboard from this thread to balance the call to OpenClipboard()
	CloseClipboard();

	// Returns 0 if file written to disk, 1 if an error occurred
	return SUCCEEDED(hr) ? 0 : 1;
}
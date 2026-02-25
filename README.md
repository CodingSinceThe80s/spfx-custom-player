# spfx-custom-player

## Summary

A SharePoint Framework (SPFx) webpart that provides a customizable HTML5 video player using the Video.js library. This solution enables users to embed and play video content directly in SharePoint pages, Microsoft Teams tabs, or other Microsoft 365 applications with full control over player settings and appearance.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js version 22.14.0 or higher (< 23.0.0)
- SharePoint Online tenant or Microsoft 365 developer tenant
- Access to video files (supports MP4, WebM, and OGG formats)
- For OneDrive/SharePoint hosted videos, ensure proper sharing permissions are configured (OneDrive support still buggy)

## Solution

| Solution           | Author(s)      |
| ------------------ | -------------- |
| spfx-custom-player | Razvan Hrestic |

## Version history

| Version | Date              | Comments        |
| ------- | ----------------- | --------------- |
| 1.0     | February 25, 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Navigate to the solution folder in your terminal
- In the command line, run:
  - `npm install`
  - `gulp serve --nobrowser`
- Open SharePoint workbench and add the "Html5CustomVideo" webpart
- Configure the video URL in the webpart property pane

To deploy:
- `gulp bundle --ship`
- `gulp package-solution --ship`
- Upload the `.sppkg` file from `sharepoint/solution` folder to your App Catalog

## Features

This webpart provides a robust HTML5 video player powered by Video.js with the following capabilities:

### Player Features
- **Multiple format support**: MP4, WebM, and OGG video formats
- **Responsive design**: Fluid layout that adapts to container size
- **Cross-origin support**: Plays videos from external sources
- **OneDrive/SharePoint integration**: Automatic URL transformation for OneDrive and SharePoint-hosted videos
- **Error handling**: Graceful error messages and loading states

### Configurable Settings
The webpart includes a comprehensive property pane with the following configuration options:

**Video Source**
- Video URL input with support for direct links and OneDrive/SharePoint URLs
- Video title/label for accessibility

**Player Settings**
- Autoplay toggle
- Player controls display toggle
- Preload options (auto, metadata, none)

**Dimensions**
- Configurable width (supports percentage and pixel values)
- Configurable height (supports pixel values)

### Technical Concepts Demonstrated
- Integration of third-party libraries (Video.js) in SPFx webparts
- React component lifecycle management with proper cleanup
- Dynamic property-based UI updates
- URL transformation for cloud storage providers
- Responsive web design patterns
- React TypeScript best practices
- SharePoint Framework property pane configuration

## Usage

1. Add the webpart to your SharePoint page
2. Click the edit icon to open the property pane
3. In the "Video Source" section, enter your video URL
   - For OneDrive/SharePoint videos, copy the sharing link
   - For external videos, use direct MP4, WebM, or OGG URLs
4. Configure optional settings:
   - Set a video title for accessibility
   - Enable/disable autoplay
   - Show/hide player controls
   - Adjust preload behavior
   - Customize player dimensions
5. Publish the page to make the video available to users

## Key Dependencies

- **Video.js** (v8.23.7): HTML5 video player library
- **React** (v17.0.1): UI component framework
- **SharePoint Framework** (v1.21.1): Microsoft 365 development platform

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Video.js Documentation](https://videojs.com/)
- [SharePoint Framework Property Pane](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/integrate-web-part-properties-with-sharepoint)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

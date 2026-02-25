import * as React from 'react';
import styles from './Html5CustomVideo.module.scss';
import type { IHtml5CustomVideoProps } from './IHtml5CustomVideoProps';
import videojs from 'video.js';
import 'video.js/dist/video-js.css';

interface VideoPlayerState {
  isLoading: boolean;
  error: string | null;
}

export default class Html5CustomVideo extends React.Component<IHtml5CustomVideoProps, VideoPlayerState> {
  private videoRef: React.RefObject<HTMLVideoElement>;
  private playerId: string;
  private playerInstance: any = null;

  constructor(props: IHtml5CustomVideoProps) {
    super(props);
    this.videoRef = React.createRef();
    this.playerId = `video-js-player-${Math.random().toString(36).substr(2, 9)}`;
    this.state = {
      isLoading: false,
      error: null
    };
  }

  public componentDidMount(): void {
    this.initializePlayer();
  }

  public componentDidUpdate(prevProps: IHtml5CustomVideoProps): void {
    if (
      prevProps.videoUrl !== this.props.videoUrl ||
      prevProps.videoTitle !== this.props.videoTitle ||
      prevProps.autoplay !== this.props.autoplay ||
      prevProps.controls !== this.props.controls ||
      prevProps.preload !== this.props.preload ||
      prevProps.width !== this.props.width ||
      prevProps.height !== this.props.height
    ) {
      this.disposePlayer();
      this.initializePlayer();
    }
  }

  public componentWillUnmount(): void {
    this.disposePlayer();
  }

  private disposePlayer = (): void => {
    if (this.playerInstance) {
      try {
        this.playerInstance.dispose();
      } catch (error) {
        console.error('Error disposing video.js player:', error);
      }
      this.playerInstance = null;
    }
  }

  private transformOneDriveUrl = (url: string): string => {
    if (url.indexOf('sharepoint.com') !== -1 || url.indexOf('1drv.ms') !== -1 || url.indexOf('onedrive.live.com') !== -1) {
      if (url.indexOf('download=1') !== -1) {
        return url;
      }
      
      if (url.indexOf('?web=1') !== -1) {
        return url.replace('?web=1', '?download=1');
      }
      
      const separator = url.indexOf('?') !== -1 ? '&' : '?';
      return url + separator + 'download=1';
    }
    return url;
  }

  private getVideoType = (url: string): string => {
    if (url.indexOf('.mp4') !== -1) {
      return 'video/mp4';
    } else if (url.indexOf('.webm') !== -1) {
      return 'video/webm';
    } else if (url.indexOf('.ogv') !== -1 || url.indexOf('.ogg') !== -1) {
      return 'video/ogg';
    }
    return 'video/mp4';
  }

  private initializePlayer = (): void => {
    const {
      videoUrl,
      videoTitle,
      autoplay,
      controls,
      preload,
      width,
      height
    } = this.props;

    if (!this.videoRef.current || !videoUrl) {
      return;
    }

    try {
      const transformedUrl = this.transformOneDriveUrl(videoUrl);
      
      const playerOptions = {
        controls: controls,
        autoplay: autoplay,
        preload: preload,
        width: width === '100%' ? undefined : width,
        height: height === '540px' ? undefined : height,
        fluid: width === '100%',
        responsive: true,
        html5: {
          hls: {
            overrideNative: true
          },
          nativeControlsForTouch: false
        },
        sources: [
          {
            src: transformedUrl,
            type: this.getVideoType(transformedUrl)
          }
        ]
      };

      if (!this.playerInstance) {
        this.playerInstance = videojs(this.videoRef.current, playerOptions);
        
        this.playerInstance.on('error', () => {
          const error = this.playerInstance.error();
          console.error('Video playback error:', error);
          this.setState({ error: 'Error loading video. Please check the URL and ensure it has proper permissions.' });
        });

        this.playerInstance.on('loadstart', () => {
          this.setState({ isLoading: true, error: null });
        });

        this.playerInstance.on('canplay', () => {
          this.setState({ isLoading: false });
        });

        if (videoTitle) {
          this.playerInstance.currentResolution = { label: videoTitle };
        }
      }
    } catch (error) {
      console.error('Error initializing video.js player:', error);
      this.setState({ error: 'Error initializing player' });
    }
  }

  public render(): React.ReactElement<IHtml5CustomVideoProps> {
    const { videoUrl, width, height, videoTitle } = this.props;
    const { isLoading, error } = this.state;

    const containerStyle: React.CSSProperties = {
      width: width || '100%',
      maxWidth: '100%'
    };

    const videoStyle: React.CSSProperties = {
      width: '100%',
      height: height || '540px'
    };

    if (!videoUrl) {
      return (
        <section className={styles.html5CustomVideo} style={containerStyle}>
          <div className={styles.welcome}>
            <h3>Video.js Player</h3>
            <p>Please configure a video URL in the webpart properties to display the player.</p>
          </div>
        </section>
      );
    }

    return (
      <section className={styles.html5CustomVideo} style={containerStyle}>
        {error && (
          <div style={{
            backgroundColor: '#fee',
            border: '1px solid #fcc',
            color: '#c33',
            padding: '12px',
            borderRadius: '4px',
            marginBottom: '10px',
            fontSize: '14px'
          }}>
            <strong>Error:</strong> {error}
          </div>
        )}
        {isLoading && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            height: height || '540px',
            backgroundColor: '#000',
            color: '#fff',
            fontSize: '14px'
          }}>
            Loading video...
          </div>
        )}
        <div
          data-vjs-player={true}
          style={{ width: '100%', height: '100%' }}
        >
          <video
            ref={this.videoRef}
            id={this.playerId}
            className="video-js vjs-default-skin"
            style={videoStyle}
            title={videoTitle}
            playsInline
            crossOrigin="anonymous"
          />
        </div>
      </section>
    );
  }
}

package example.hamid.com.videoplayer;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.widget.VideoView;
import android.widget.MediaController;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        final VideoView videoView=(VideoView)findViewById(R.id.videoView);
                videoView.setVideoPath("http://geeknation.com/videos/star-wars-battlefront-reveal-trailer-2015/");
        videoView.start();


    }


}

package com.salty.payslip.Activity

import android.content.Intent
import android.os.Bundle
import android.os.Handler
import android.os.Looper
import android.view.animation.AnimationUtils
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import com.salty.payslip.R
class SplashActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_splash2)

        val ivLogo = findViewById<android.widget.ImageView>(R.id.ivLogo)
        val tvAppName = findViewById<android.widget.TextView>(R.id.tvAppName)
        val tvPowered = findViewById<android.widget.TextView>(R.id.tvPowered)

        // Logo: Zoom + Fade + Slight Rotate
        val logoAnim = AnimationUtils.loadAnimation(this, R.anim.zoom_fade_in)
        ivLogo.startAnimation(logoAnim)

        // App Name: Slide up + Fade
        val textAnim = AnimationUtils.loadAnimation(this, R.anim.slide_up_fade)
        textAnim.startOffset = 600  // Start after logo begins
        tvAppName.startAnimation(textAnim)

        val poweredAnim = AnimationUtils.loadAnimation(this, R.anim.fade_in_slow)
        tvPowered.startAnimation(poweredAnim)

        // Go to ThirdActivity after full animation
        Handler(Looper.getMainLooper()).postDelayed({
            startActivity(Intent(this, ThirdActivity::class.java))
            overridePendingTransition(android.R.anim.fade_in, android.R.anim.fade_out)
            finish()
        }, 2800)
    }
}
package utils;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class ApplicationRrun {
    public static void main(String[] args) {
        SpringApplication.run(ApplicationRrun.class,args);
    }
}

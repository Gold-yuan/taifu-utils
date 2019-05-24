package sichuan.ytf.thread;

import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import com.google.common.util.concurrent.SimpleTimeLimiter;
import com.google.common.util.concurrent.TimeLimiter;

public class ThreadTest {
    public static void main(String[] args) {
        System.out.println("main begin");
        
        Thread thread = new Thread(new Runnable() {
            @Override
            public void run() {
                System.out.println("thread begin");
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                System.out.println("thread end");
                
            }
            @Override
            protected void finalize() throws Throwable {
                System.out.println("finally");
            }
            
        });
        thread.setDaemon(true);
        thread.start();
        System.out.println("main end");
    }

    public static void a() {

        TimeLimiter timeLimiter = SimpleTimeLimiter.create(Executors.newFixedThreadPool(20));
        UserInfoService userInfoService = new UserInfoServiceImpl();
        UserInfoService userInfoService1 = timeLimiter.newProxy(userInfoService, UserInfoService.class, 2000,
                TimeUnit.MILLISECONDS);
        try {
            String userName = userInfoService1.getUserName();
            System.out.println("userName:" + userName);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static class UserInfoServiceImpl implements UserInfoService {

        @Override
        public String getUserName() throws InterruptedException {
            System.out.println("123");
            Thread.sleep(5000);
            System.out.println("456");
            return "tim";
        }
    }
}

interface UserInfoService {
    String getUserName() throws InterruptedException;
}

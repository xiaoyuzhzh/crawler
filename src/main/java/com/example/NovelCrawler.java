package com.example;

import java.nio.charset.Charset;

import cn.edu.hfut.dmic.webcollector.model.CrawlDatum;
import cn.edu.hfut.dmic.webcollector.model.CrawlDatums;
import cn.edu.hfut.dmic.webcollector.model.Page;
import cn.edu.hfut.dmic.webcollector.net.HttpRequest;
import cn.edu.hfut.dmic.webcollector.net.HttpResponse;
import cn.edu.hfut.dmic.webcollector.plugin.berkeley.BreadthCrawler;

/**
 * Created by xiaoy on 2017/4/19.
 */
public class NovelCrawler extends BreadthCrawler {

    Charset charset = Charset.forName("utf-8");

    public NovelCrawler(String crawlPath, boolean autoParse) {
        super(crawlPath, autoParse);
        /*start page*/
        this.addSeed("https://api.sfacg.com/ajax/ashx/Common.ashx?op=getFansList&avatar=true&nid=40141&_=1489570443947");
        this.addSeed("https://api.sfacg.com/novels?size=8&filter=newpush");




//    /*fetch url like http://news.hfut.edu.cn/show-xxxxxxhtml*/
//    this.addRegex("http://news.hfut.edu.cn/show-.*html");
//    /*do not fetch jpg|png|gif*/
//    this.addRegex("-.*\\.(jpg|png|gif).*");
//    /*do not fetch url contains #*/
//    this.addRegex("-.*#.*");
    }

    @Override
    public void visit(Page page, CrawlDatums next) {
        String url = page.getUrl();
        System.out.println("url:"+url);
//        if(page.getResponse().getHeaders().get("Content-Type").get(0).equals("text/html; charset=utf-8")){
            System.out.println("url: "+url+";  content: "+new String(page.getResponse().getContent(),charset));
//        }
    }

    @Override
    public HttpResponse getResponse(CrawlDatum crawlDatum) throws Exception {
        HttpRequest request = new HttpRequest(crawlDatum.getUrl());

        request.setMethod("GET");
        String outputData=crawlDatum.getMetaData("outputData");
        if(outputData!=null){
            request.setOutputData(outputData.getBytes("utf-8"));
        }
        request.setHeader("Authorization","Basic YW5kcm9pZHVzZXI6MWEjJDUxLXl0Njk7KkFjdkBxeHE=");
        return request.getResponse();
        /*
        //通过下面方式可以设置Cookie、User-Agent等http请求头信息
        request.setCookie("xxxxxxxxxxxxxx");
        request.setUserAgent("WebCollector");
        request.addHeader("xxx", "xxxxxxxxx");
        */
    }

    public static void main(String[] args) throws Exception {
        NovelCrawler crawler = new NovelCrawler("crawl", true);
        crawler.setThreads(50);
        crawler.setTopN(100);
        //crawler.setResumable(true);
    /*start crawl with depth of 4*/
        crawler.start(4);
    }
}

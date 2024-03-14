
jQuery(document).ready(function () {

    /*
    Background slideshow
    */
    $.backstretch([
      "assets/img/backgrounds/1.jpg"
    , "assets/img/backgrounds/2.jpg"
    , "assets/img/backgrounds/3.jpg"
    ], { duration: 3000, fade: 750 });

    /*
    Tooltips
    */
    $('.links a.home').tooltip();
    $('.links a.blog').tooltip();

    /*
    Form validation
    
    $('.register form').submit(function () {
        $(this).find("label[for='tb_Name']").html('濮撳悕');
        $(this).find("label[for='tb_Tel']").html('鑱旂郴鏂瑰紡');
        $(this).find("label[for='username']").html('鐢ㄦ埛鍚�');
        $(this).find("label[for='tb_Email']").html('閭');
        $(this).find("label[for='password']").html('瀵嗙爜');

        $(this).find("label[for='tb_Title']").html('鎰忚涓绘棬');
        $(this).find("label[for='tb_Content']").html('鍐呭鍙欒堪');

        ////
        var tb_Name = $(this).find('input#tb_Name').val();
        var tb_Tel = $(this).find('input#tb_Tel').val();
        var username = $(this).find('input#username').val();
        var tb_Email = $(this).find('input#tb_Email').val();
        var password = $(this).find('input#password').val();
   

        var tb_Title = $(this).find('input#tb_Title').val();
        var tb_Content = $(this).find('input#tb_Content').val();


        if (tb_Name == '') {
            $(this).find("label[for='tb_Name']").append("<span style='display:none' class='red'> - 瀹堕暱鎮ㄥソ锛岃杈撳叆濮撳悕銆�</span>");
            $(this).find("label[for='tb_Name'] span").fadeIn('medium');
            return false;
        }
        if (tb_Tel == '') {
            $(this).find("label[for='tb_Tel']").append("<span style='display:none' class='red'> - 璇疯緭鍏ヨ仈绯荤數璇濄€�</span>");
            $(this).find("label[for='tb_Tel'] span").fadeIn('medium');
            return false;
        }
        if (username == '') {
            $(this).find("label[for='username']").append("<span style='display:none' class='red'> - 璇疯緭鍏ュ鐢熷鍙枫€�</span>");
            $(this).find("label[for='username'] span").fadeIn('medium');
            return false;
        }
        if (tb_Email == '') {
            $(this).find("label[for='tb_Email']").append("<span style='display:none' class='red'> - 璇疯緭鍏ラ偖绠便€�</span>");
            $(this).find("label[for='tb_Email'] span").fadeIn('medium');
            return false;
        }
        if (password == '') {
            $(this).find("label[for='password']").append("<span style='display:none' class='red'> - 璇疯緭鍏ュ瘑鐮併€�</span>");
            $(this).find("label[for='password'] span").fadeIn('medium');
            return false;
        }

        if (tb_Title == '') {
            $(this).find("label[for='tb_Title']").append("<span style='display:none' class='red'> - 璇疯緭鍏ユ剰瑙佷富鏃ㄣ€�</span>");
            $(this).find("label[for='tb_Title'] span").fadeIn('medium');
            return false;
        }
        if (tb_Content == '') {
            $(this).find("label[for='tb_Content']").append("<span style='display:none' class='red'> - 璇疯緭鍏ュ唴瀹瑰彊杩般€�</span>");
            $(this).find("label[for='tb_Content'] span").fadeIn('medium');
            return false;
        }
    });
    */
});



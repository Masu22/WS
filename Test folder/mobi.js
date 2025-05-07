// サイドメニューの表示
document.addEventListener("DOMContentLoaded", function() {
    const menuButton = document.getElementById("menu-button");
    const sideMenu = document.getElementById("side-menu");
    const closeButton = document.getElementById("close-button");

    // メニューを開く
    menuButton.addEventListener("click", function() {
        sideMenu.classList.add("open");
    });

    // メニューを閉じる
    closeButton.addEventListener("click", function() {
        sideMenu.classList.remove("open");
    });

    // メニューの外側をクリックしたら閉じる
    document.addEventListener("click", function(event) {
        if (!sideMenu.contains(event.target) && event.target !== menuButton) {
            sideMenu.classList.remove("open");
        }
    });
});


// フッターメニューの表示
document.addEventListener("DOMContentLoaded", function() {
    const footerMenu = document.getElementById("footer-menu");
    let lastScrollTop = 0;

    window.addEventListener("scroll", function() {
        let scrollTop = window.scrollY || document.documentElement.scrollTop;

        if (scrollTop > 50) { // 50px以上スクロールしたら表示
            footerMenu.classList.add("show");
        } else {
            footerMenu.classList.remove("show"); // 一番上では隠す
        }

        lastScrollTop = scrollTop;
    });
});
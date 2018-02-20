$(function() {
    "use strict";
    var e = (s = $(window)).scrollTop(),
        a = $(".nav-wrapper");
    $('a[title="Hosted on free web hosting 000webhost.com. Host your own website for FREE."]').remove();
    var s = $(window),
        o = $(".loader-con");
    s.on("load", function() {
        o.fadeOut("slow"), setTimeout(function() {
            o.css("display", "none")
        }, 600), s.scrollTop(0)
    });
    var t = !1,
        l = "Webkit Moz O ms Khtml".split(" "),
        r = "",
        c = document.createElement("div");
    if (void 0 !== c.style.animationName && (t = !0), !1 === t)
        for (var i = 0; i < l.length; i++)
            if (void 0 !== c.style[l[i] + "AnimationName"]) {
                (r = l[i]) + "Animation", "-" + r.toLowerCase() + "-", t = !0;
                break
            }
    $('a.scroll[href*="#"]').not('[href="#"]').not('[href="#0"]').on("click", function(e) {
        if (location.pathname.replace(/^\//, "") == this.pathname.replace(/^\//, "") && location.hostname == this.hostname) {
            var a = $(this.hash),
                s = $(this).data("speed") || 800;
            (a = a.length ? a : $("[name=" + this.hash.slice(1) + "]")).length && (e.preventDefault(), $("html, body").animate({
                scrollTop: a.offset().top
            }, s))
        }
    }), $(".scroll-up").on("click", function(e) {
        e.preventDefault(), $("html, body").animate({
            scrollTop: 0
        }, 900)
    });
    var n = a.find(".navbar-collapse ul li a");
    $.each(n, function(e, s) {
        $(this).on("click", function(e) {
            a.find(".navbar-collapse").collapse("hide")
        })
    });
    var h = $(".one-page-section"),
        p = a.height();
    var m = $(".parallax");
    if (t) {
        var f = $("#home .typer-title");
        $.prototype.typer && f.typer(["I'm Software Developer", "Let's Work Together", "Let's Create Awesome Stuff"])
    }
    var d = $(".animate-on-scroll");
    var u = $(".timer-con");
    s.on("scroll", function() {
        (e = s.scrollTop()) >= 100 ? a.addClass("nav-sticky") : a.removeClass("nav-sticky"), e >= 1e3 ? $(".scroll-up").addClass("show-up-btn") : $(".scroll-up").removeClass("show-up-btn"), e = s.scrollTop(), $.each(h, function(s, o) {
                var t = $(this),
                    l = t.offset().top - p,
                    r = l + t.height();
                if (e >= l && e <= r) {
                    var c = a.find("ul li a");
                    $.each(c, function(e, a) {
                        $(this).removeClass("active")
                    }), a.find('ul a[href="#' + t.attr("id") + '"]').addClass("active")
                }
            }), $.each(m, function(a, o) {
                var t = $(this),
                    l = t.offset().top,
                    r = t.height(),
                    c = ((e = s.scrollTop()) - l) / 2;
                c <= r + l && t.css({
                    backgroundPosition: "50% " + c + "px"
                })
            }),
            function() {
                if (u.length && !u.hasClass("done")) {
                    var e = u.offset().top;
                    s.scrollTop() >= e - (s.height() - 100) && (u.addClass("done"), $.prototype.countTo && $(".timer").countTo({
                        speed: 2500,
                        refreshInterval: 50
                    }))
                }
            }(), $.each(d, function(e, a) {
                var o = $(this),
                    l = o.data("animation"),
                    r = o.data("delay") || .2;
                if (t && !o.hasClass(l)) {
                    o.css("opacity", 0);
                    var c = o.offset().top;
                    s.scrollTop() >= c - (s.height() - 100) && (o.css("animationDelay", r + "s"), o.addClass(l), setTimeout(function() {
                        o.css("opacity", 1)
                    }, 1e3))
                }
            })
    });
    var g = $(".skills-init"),
        b = $(".hide-skill-bar");
    g.on("shown.bs.tab", function(e) {
        g.hasClass("done") || (g.addClass("done"), $.each(b, function(e, a) {
            var s = $(this),
                o = s.find(".progress-bar-text"),
                t = parseInt(o.data("percent"), 10),
                l = $(s).closest(".progress-bar-skills").find("p");
            s.css("height", 100 - t + "%"), l.html(t + "%")
        }))
    });
    var v = $(".work-item"),
        y = $("#work-list"),
        C = $(".filter"),
        k = function(e) {
            var a = $(this);
            if (e.preventDefault(), a.hasClass("filter-active")) return !1;
            y.find(".filter-active").removeClass("filter-active"), a.addClass("filter-active"), $.each(v, function(e, s) {
                var o = $(this);
                "all" === a.data("filter") ? (o.removeClass("filtered"), setTimeout(function() {
                    o.css("display", "block")
                }, 500)) : o.hasClass(a.data("filter")) ? (o.removeClass("filtered"), setTimeout(function() {
                    o.css("display", "block")
                }, 500)) : (o.addClass("filtered"), setTimeout(function() {
                    o.css("display", "none")
                }, 500))
            })
        };

    function w(e) {
        return (e = e.match(/^rgba?[\s+]?\([\s+]?(\d+)[\s+]?,[\s+]?(\d+)[\s+]?,[\s+]?(\d+)[\s+]?/i)) && 4 === e.length ? "#" + ("0" + parseInt(e[1], 10).toString(16)).slice(-2) + ("0" + parseInt(e[2], 10).toString(16)).slice(-2) + ("0" + parseInt(e[3], 10).toString(16)).slice(-2) : ""
    }
    if ($.each(C, function(e, a) {
            $(this).on("click", k)
        }), "undefined" != typeof particlesJS) {
        var T = w($(".main-color").css("color"));
        particlesJS("particles-js", {
            particles: {
                number: {
                    value: 80,
                    density: {
                        enable: !0,
                        value_area: 600
                    }
                },
                color: {
                    value: T
                },
                shape: {
                    type: "circle",
                    stroke: {
                        width: 0,
                        color: "#888"
                    },
                    polygon: {
                        nb_sides: 5
                    },
                    image: {
                        src: "img/github.svg",
                        width: 100,
                        height: 100
                    }
                },
                opacity: {
                    value: .6,
                    random: !1,
                    anim: {
                        enable: !1,
                        speed: 1,
                        opacity_min: .1,
                        sync: !1
                    }
                },
                size: {
                    value: 3,
                    random: !0,
                    anim: {
                        enable: !1,
                        speed: 40,
                        size_min: .1,
                        sync: !1
                    }
                },
                line_linked: {
                    enable: !0,
                    distance: 150,
                    color: "#888",
                    opacity: .4,
                    width: 1
                },
                move: {
                    enable: !0,
                    speed: 5,
                    direction: "bottom",
                    random: !1,
                    straight: !1,
                    out_mode: "out",
                    attract: {
                        enable: !1,
                        rotateX: 600,
                        rotateY: 1200
                    }
                }
            },
            interactivity: {
                detect_on: "canvas",
                events: {
                    onhover: {
                        enable: !0,
                        mode: "repulse"
                    },
                    onclick: {
                        enable: !0,
                        mode: "push"
                    },
                    resize: !0
                },
                modes: {
                    grab: {
                        distance: 400,
                        line_linked: {
                            opacity: 1
                        }
                    },
                    bubble: {
                        distance: 400,
                        size: 30,
                        duration: 2,
                        opacity: 8,
                        speed: 3
                    },
                    repulse: {
                        distance: 200
                    },
                    push: {
                        particles_nb: 4
                    },
                    remove: {
                        particles_nb: 2
                    }
                }
            },
            retina_detect: !0,
            config_demo: {
                hide_card: !1,
                background_color: "#b61924",
                background_image: "",
                background_position: "50% 50%",
                background_repeat: "no-repeat",
                background_size: "cover"
            }
        })
    }
    if ($.prototype.magnificPopup && ($(".popup-link").magnificPopup({
            type: "image",
            removalDelay: 300,
            mainClass: "mfp-fade",
            gallery: {
                enabled: !0
            }
        }), $(".pf-gallery-popuap").magnificPopup({
            type: "image",
            removalDelay: 300,
            mainClass: "mfp-fade",
            gallery: {
                enabled: !0
            },
            zoom: {
                enabled: !0,
                duration: 300,
                easing: "ease-in-out"
            }
        })), $.prototype.owlCarousel) {
        function _(e) {
            var a = $(".main-color").css("color"),
                s = e.find(".owl-dot");
            $.each(s, function(e, s) {
                $(this).hasClass("active") ? $(this).css("background-color", a) : $(this).css("background-color", "#fff")
            })
        }
        var x = $(".home-3 .home-carousel");
        x.owlCarousel({
            mouseDrag: !0,
            nav: !1,
            dots: !1,
            items: 1,
            autoplay: !0,
            autoplayHoverPause: !0,
            autoplayTimeout: 3e3,
            loop: !0,
            animateOut: "fade-out",
            animateIn: "slide-animation",
            responsive: {
                768: {
                    dots: !0
                }
            }
        }), _(x), x.on("changed.owl.carousel", function(e) {
            _(x)
        });
        var I = $(".testimonials-carousel");
        I.owlCarousel({
            dots: !0,
            items: 1,
            loop: !0,
            autoplay: !0,
            autoplayHoverPause: !0,
            autoplayTimeout: 3e3
        }), _(I), I.on("changed.owl.carousel", function(e) {
            _(I)
        }), $(".pf-details-slider").owlCarousel({
            nav: !0,
            navText: ["<a class='main-color-bg pf-slider-btn ver-center'><i class='fa fa-chevron-left center effect'></i><span></span></a>", "<a class='main-color-bg pf-slider-btn ver-center'><i class='fa fa-chevron-right center effect'></i><span></span></a>"],
            items: 1,
            loop: !0,
            dots: !1
        })
    }
    var z = $("#contact-form"),
        D = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{3,4})$/,
        A = $(".input-field");

    function S() {
        if ($(this).is("#email")) {
            var e = $(this).val();
            if (!D.test(e)) return $(".email-error").html("please enter a valid email."), !1;
            $(".email-error").html("")
        } else {
            var a = $(this).attr("id"),
                s = $("." + a + "-error");
            if ("" === $(this).val()) return s.html("please enter a valid " + a + "."), !1;
            s.html(" ")
        }
    }
    $.each(A, function(e, a) {
        $(this).on("blur", S)
    });
    var P = $("#form-message");

    function L(e, a, s) {
        e /= 255, a /= 255, s /= 255;
        var o, t = Math.max(e, a, s),
            l = Math.min(e, a, s);
        if (t == l) o = 0;
        else {
            var r = t - l;
            switch ((t + l) / 2 > .5 ? r / (2 - t - l) : r / (t + l), t) {
                case e:
                    o = (a - s) / r + (a < s ? 6 : 0);
                    break;
                case a:
                    o = (s - e) / r + 2;
                    break;
                case s:
                    o = (e - a) / r + 4
            }
            o /= 6
        }
        return 360 * o
    }
    $(z).on("submit", function(e) {
        e.preventDefault();
        var a = $(z).serialize();
        $.ajax({
            type: "POST",
            url: z.attr("action"),
            data: a
        }).done(function(e) {
            P.removeClass("error"), P.addClass("success"), P.text(e), $("#name").val(""), $("#email").val(""), $("#message").val("")
        }).fail(function(e) {
            P.removeClass("success"), P.addClass("error"), "" !== e.responseText ? P.text(e.responseText) : P.text("Sorry! An error occured and your message could not be sent.")
        })
    }), $("#color-scheme-btn").click(function(e) {
        e.preventDefault(), $(".color-scheme-panel-con").toggleClass("color-scheme-panel-open ")
    });
    var O, E, H = $(".change-color"),
        M = $(".main-color"),
        Z = $(".main-color-bg"),
        j = $(".section-title img"),
        J = w(M.css("color")),
        N = parseInt(J.substr(1, 2), 16),
        R = parseInt(J.substr(3, 2), 16),
        U = parseInt(J.substr(5, 2), 16);
    $.each(j, function(e, a) {
        $(this).css("filter", "hue-rotate(" + L(N, R, U) + "deg)")
    }), $.each(H, function(e, a) {
        $(this).on("click", function(e) {
            e.preventDefault();
            var a = $(this).attr("class").split(" "),
                s = a[0],
                o = a[1];
            $.each(M, function(e, a) {
                $(this).hasClass("main-color") ? ($(this).removeClass("main-color"), $(this).addClass(o)) : $(this).hasClass("change-color") || ($(this).removeClass(O), $(this).addClass(o))
            }), $.each(Z, function(e, a) {
                $(this).hasClass("main-color-bg") ? ($(this).removeClass("main-color-bg"), $(this).addClass(s)) : $(this).hasClass("change-color") || ($(this).removeClass(E), $(this).addClass(s))
            }), M = $("." + o), Z = $("." + s), O = o, E = s, J = w($("." + o).css("color")), N = parseInt(J.substr(1, 2), 16), R = parseInt(J.substr(3, 2), 16), U = parseInt(J.substr(5, 2), 16), $.each(j, function(e, a) {
                $(this).css("filter", "hue-rotate(" + L(N, R, U) + "deg)")
            }), $(".navbar-brand>img").css("filter", "hue-rotate(" + L(N, R, U) + "deg)")
        })
    })
});
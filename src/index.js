import "./scss/main.scss";
import "bootstrap";
import imagesLoaded from "imagesloaded";
import Masonry from "masonry-layout";

window.addEventListener("load", () => {
  var overview = document.querySelector("#overview");
  var msnry;

  imagesLoaded(overview, () => {
    msnry = new Masonry(overview, {
      itemSelector: ".grid-item",
      columnWidth: ".grid-item",
      percentPosition: true,
    });
    msnry.layout();
  });
});

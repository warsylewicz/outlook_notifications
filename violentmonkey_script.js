// ==UserScript==
// @name        New script - office.com
// @namespace   Violentmonkey Scripts
// @match       https://outlook.office.com/mail/*
// @match       https://outlook.office365.com/mail/*
// @grant       none
// @version     1.0
// @author      -
// @description 1/8/2020, 4:31:56 PM
// ==/UserScript==


let unreadEmailsCount = 0;
startMonitor();


function createIconElement() {
  let icon = document.createElement("link");
  icon.rel = "icon";
  icon.type = "image/png";
  icon.sizes = "16x16";
  icon.href = generateTabIcon(0);
  return icon;
}


function generateTabIcon(number) {
  const canvas = document.createElement("canvas");
  canvas.width = 16;
  canvas.height = 16;
  const ctx = canvas.getContext("2d");
  // draw cliped circle (radius is bigger than canvas by 1px) - the effect is a
  // square with round corners
  
  // get same background colour
  ctx.fillStyle = "#0099FF";
  let bc = document.querySelector('.o365sx-navbar');
  if (bc != null) {
    ctx.fillStyle = window.getComputedStyle(bc).getPropertyValue('background-color');
  } else {
    ctx.fillStyle = "#0099FF";
  }
  ctx.beginPath();
  ctx.arc(8, 8, 9, 0, 2 * Math.PI);
  ctx.fill();
  // draw the number in center
  ctx.font = "bold 12px Helvetica, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = "white";
  ctx.shadowBlur = 1;
  ctx.shadowColor = "black";
  ctx.shadowOffsetX = 1;
  ctx.shadowOffsetY = 1;
  ctx.fillText(number, 8, 8);
  return canvas.toDataURL("image/png");
}


function setFavicon(count) {
  let nr = Math.min(count, 99);
  
  if (nr === 0) {
    nr = "-";
  }
  
  let owaIcon = createIconElement();
  
  owaIcon.href = generateTabIcon(nr);
  document.head.appendChild(owaIcon);
}


function extractNumber(text) {
  if (text) {
    let digits = text.match(/\d/gi);
    if (digits) {
      return parseInt(digits.join(""), 10);
    }
  }
  return 0;
}

function getCountFromNodes(nodes) {
  let count = 0;
  if (nodes) {
    for (let i = nodes.length - 1; i >= 0; i--) {
      count += extractNumber(nodes[i].innerHTML);
    }
  }
  return count;
}

function singularOrPlural(word, count) {
  return word + ((count === 1) ? "" : "s");
}


function triggerNotification(type, text) {
  if (!prefs.disableNotifications) {
    browser.runtime.sendMessage({
      "type" : type,
      "msg" : text
    });
  }
}

function checkForNewMessages() {
  let newUnreadEmailsCount = getCountFromNodes(document.querySelectorAll("div[title=\"Inbox\"] > span:nth-child(3) > span"));
  if (newUnreadEmailsCount !== unreadEmailsCount && newUnreadEmailsCount > 0) {
    let notify = new Notification('You have ' + newUnreadEmailsCount + ' new ' + singularOrPlural("email", newUnreadEmailsCount) + '.');
  }
  
  unreadEmailsCount = newUnreadEmailsCount;
  setFavicon(unreadEmailsCount);
}


function startMonitor() {
  setFavicon(0);
  checkForNewMessages();
  newEventsTimer = setInterval(checkForNewMessages, 5000);
}



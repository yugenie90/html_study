const categories = {
    korean: ["불고기", "비빔밥", "김치찌개", "삼겹살", "떡볶이", "갈비찜", "잡채", "순두부찌개", "제육볶음", "냉면"],
    japanese: ["초밥", "라멘", "돈까스", "우동", "타코야끼", "사시미", "야끼니꾸", "오코노미야끼", "규동", "모밀"],
    chinese: ["짜장면", "짬뽕", "탕수육", "마라탕", "꿔바로우", "깐풍기", "양꼬치", "유린기", "마파두부", "볶음밥"],
    western: ["파스타", "피자", "스테이크", "햄버거", "샐러드", "리조또", "감바스", "뇨끼", "라자냐", "타코"]
};
categories.all = [...new Set([...categories.korean, ...categories.japanese, ...categories.chinese, ...categories.western])];

let foods = [...categories.all];
const colors = ["#FFC0CB", "#FFD700", "#FF69B4", "#ADD8E6", "#90EE90", "#FFA07A", "#20B2AA", "#87CEFA", "#778899", "#B0C4DE", "#FFB6C1", "#F0E68C", "#E6E6FA", "#FAFAD2", "#D3FFCE", "#FFE4E1", "#AFEEEE", "#DB7093", "#F5DEB3", "#FFFFFF"];

const canvas = document.getElementById('roulette-canvas');
const ctx = canvas.getContext('2d');
const spinButton = document.getElementById('spin-button');
const resultDiv = document.getElementById('result');
const categoryButtons = document.querySelectorAll('.category-button');
const categoryResultDiv = document.getElementById('category-result');

let currentAngle = 0;
let spinAngleStart = 0;
let spinTime = 0;
let spinTimeTotal = 0;

function drawRoulette() {
    if (!canvas.getContext) return;
    const arc = Math.PI / (foods.length / 2);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.strokeStyle = "#333";
    ctx.lineWidth = 2;
    ctx.font = '20px "Dongle", sans-serif';

    for (let i = 0; i < foods.length; i++) {
        const angle = currentAngle + i * arc;
        ctx.fillStyle = colors[i % colors.length];

        ctx.beginPath();
        ctx.arc(200, 200, 200, angle, angle + arc, false);
        ctx.arc(200, 200, 0, angle + arc, angle, true);
        ctx.stroke();
        ctx.fill();

        ctx.save();
        ctx.fillStyle = "black";
        ctx.translate(200 + Math.cos(angle + arc / 2) * 150, 200 + Math.sin(angle + arc / 2) * 150);
        ctx.rotate(angle + arc / 2 + Math.PI / 2);
        const text = foods[i];
        ctx.fillText(text, -ctx.measureText(text).width / 2, 0);
        ctx.restore();
    }
}

function spin() {
    resultDiv.innerHTML = '';
    spinAngleStart = Math.random() * 10 + 10;
    spinTime = 0;
    spinTimeTotal = Math.random() * 3 + 4 * 1000;
    rotateRoulette();
}

function rotateRoulette() {
    spinTime += 30;
    if (spinTime >= spinTimeTotal) {
        stopRotateRoulette();
        return;
    }
    const spinAngle = spinAngleStart - easeOut(spinTime, 0, spinAngleStart, spinTimeTotal);
    currentAngle += (spinAngle * Math.PI / 180);
    drawRoulette();
    requestAnimationFrame(rotateRoulette);
}

function stopRotateRoulette() {
    const totalArcs = foods.length;
    const arcSize = 2 * Math.PI / totalArcs;
    const finalAngle = currentAngle + (Math.PI / 2); // Adjust for top pointer
    let winningIndex = Math.floor((2 * Math.PI - (finalAngle % (2*Math.PI)))) / arcSize;
    winningIndex = Math.floor(winningIndex) % totalArcs;

    const text = foods[winningIndex];
    resultDiv.innerHTML = `오늘의 메뉴는... ${text}!`;
}

function easeOut(t, b, c, d) {
    const ts = (t /= d) * t;
    const tc = ts * t;
    return b + c * (tc + -3 * ts + 3 * t);
}

spinButton.addEventListener('click', spin);

categoryButtons.forEach(button => {
    button.addEventListener('click', () => {
        // Update active button
        categoryButtons.forEach(btn => btn.classList.remove('active'));
        button.classList.add('active');

        // Update roulette foods
        const category = button.dataset.category;
        foods = [...categories[category]];
        drawRoulette();
        
        // Also show a random recommendation from the category
        const randomFood = categories[category][Math.floor(Math.random() * categories[category].length)];
        categoryResultDiv.innerHTML = `추천 메뉴: <span style="color: #e64980;">${randomFood}</span>`;
    });
});

drawRoulette();
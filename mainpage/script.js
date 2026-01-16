// Mobile menu toggle
const menuToggle = document.getElementById('menuToggle');
const navLinks = document.getElementById('navLinks');

if (menuToggle) {
    menuToggle.addEventListener('click', () => {
        navLinks.classList.toggle('active');
        menuToggle.textContent = navLinks.classList.contains('active') ? '✕' : '☰';
    });

    // Close menu when clicking on a link
    document.querySelectorAll('.nav-link, .btn-nav-download').forEach(link => {
        link.addEventListener('click', () => {
            navLinks.classList.remove('active');
            menuToggle.textContent = '☰';
        });
    });

    // Close menu when clicking outside
    document.addEventListener('click', (e) => {
        if (!navLinks.contains(e.target) && !menuToggle.contains(e.target)) {
            navLinks.classList.remove('active');
            menuToggle.textContent = '☰';
        }
    });
}

// Smooth scroll for navigation links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});

// Add active state to nav on scroll
const navbar = document.querySelector('.navbar');
window.addEventListener('scroll', () => {
    if (window.scrollY > 50) {
        navbar.style.boxShadow = '0 4px 12px rgba(0, 0, 0, 0.1)';
    } else {
        navbar.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.1)';
    }
});

// Animate elements on scroll
const observerOptions = {
    threshold: 0.1,
    rootMargin: '0px 0px -50px 0px'
};

const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
        if (entry.isIntersecting) {
            entry.target.style.opacity = '1';
            entry.target.style.transform = 'translateY(0)';
        }
    });
}, observerOptions);

// Observe feature cards and other elements
document.addEventListener('DOMContentLoaded', () => {
    const animatedElements = document.querySelectorAll('.feature-card, .download-card, .why-item');

    animatedElements.forEach(el => {
        el.style.opacity = '0';
        el.style.transform = 'translateY(20px)';
        el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(el);
    });
});

// Detect OS and highlight appropriate download button
function detectOS() {
    const userAgent = window.navigator.userAgent.toLowerCase();
    const platform = window.navigator.platform.toLowerCase();

    let os = 'windows'; // default

    if (platform.includes('mac')) {
        os = 'mac';
    } else if (platform.includes('linux') || platform.includes('x11')) {
        os = 'linux';
    } else if (userAgent.includes('win')) {
        os = 'windows';
    }

    return os;
}

// Update hero download button based on OS
document.addEventListener('DOMContentLoaded', () => {
    const os = detectOS();
    const heroDownloadBtn = document.querySelector('.hero-buttons .btn-primary');

    if (heroDownloadBtn) {
        const iconSpan = heroDownloadBtn.querySelector('.btn-icon');
        const textNode = Array.from(heroDownloadBtn.childNodes).find(node => node.nodeType === 3);

        switch (os) {
            case 'linux':
                if (textNode) textNode.textContent = 'Download for Linux';
                break;
            case 'mac':
                if (textNode) textNode.textContent = 'Download for Mac';
                break;
            default:
                // Keep Windows as default
                break;
        }
    }
});

// Add hover effect to preview items
document.addEventListener('DOMContentLoaded', () => {
    const previewItems = document.querySelectorAll('.preview-item');
    let currentIndex = 0;

    setInterval(() => {
        previewItems.forEach(item => item.classList.remove('active'));
        currentIndex = (currentIndex + 1) % previewItems.length;
        previewItems[currentIndex].classList.add('active');
    }, 2000);
});

// Mobile menu toggle (if you want to add mobile menu later)
const createMobileMenu = () => {
    const nav = document.querySelector('.nav-links');
    const burger = document.createElement('button');
    burger.className = 'mobile-menu-toggle';
    burger.innerHTML = '☰';
    burger.style.display = 'none';

    if (window.innerWidth <= 640) {
        burger.style.display = 'block';
    }

    burger.addEventListener('click', () => {
        nav.classList.toggle('mobile-menu-open');
    });

    document.querySelector('.nav-container').appendChild(burger);
};

// Typing effect for hero title (optional enhancement)
function typeWriter(element, text, speed = 100) {
    let i = 0;
    element.textContent = '';

    function type() {
        if (i < text.length) {
            element.textContent += text.charAt(i);
            i++;
            setTimeout(type, speed);
        }
    }

    type();
}

// Add download tracking (optional - for analytics)
document.querySelectorAll('.btn-download, .btn-primary').forEach(button => {
    button.addEventListener('click', (e) => {
        const buttonText = e.target.textContent.trim();
        console.log('Download button clicked:', buttonText);
        // Add your analytics tracking here
        // Example: gtag('event', 'download', { platform: 'windows' });
    });
});

const currentPage = window.location.pathname.split('/').pop() || 'index.html';
const navLinks = document.querySelectorAll('.site-nav a');

const THEME_KEY = 'ifs-theme';

function getPreferredTheme() {
    const savedTheme = localStorage.getItem(THEME_KEY);
    if (savedTheme === 'light' || savedTheme === 'dark') {
        return savedTheme;
    }
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark';
}

function applyTheme(theme, toggleButton) {
    const isLight = theme === 'light';
    document.body.classList.toggle('light-mode', isLight);
    if (toggleButton) {
        toggleButton.textContent = isLight ? 'Dark' : 'Light';
        toggleButton.setAttribute('aria-label', isLight ? 'Switch to dark mode' : 'Switch to light mode');
        toggleButton.setAttribute('aria-pressed', String(isLight));
    }
}

function setupThemeToggle() {
    const headerInner = document.querySelector('.header-inner');
    if (!headerInner) {
        return;
    }

    const themeButton = document.createElement('button');
    themeButton.type = 'button';
    themeButton.className = 'theme-toggle';

    const projectId = headerInner.querySelector('.project-id');
    if (projectId) {
        headerInner.insertBefore(themeButton, projectId);
    } else {
        headerInner.appendChild(themeButton);
    }

    let currentTheme = getPreferredTheme();
    applyTheme(currentTheme, themeButton);

    themeButton.addEventListener('click', () => {
        currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
        localStorage.setItem(THEME_KEY, currentTheme);
        applyTheme(currentTheme, themeButton);
    });
}

setupThemeToggle();

navLinks.forEach((link) => {
    const href = link.getAttribute('href');
    if (href === currentPage) {
        link.classList.add('active');
    }
});

const navToggle = document.querySelector('.nav-toggle');
const siteNav = document.querySelector('.site-nav');

if (navToggle && siteNav) {
    navToggle.addEventListener('click', () => {
        siteNav.classList.toggle('open');
    });
}

const revealItems = document.querySelectorAll('.reveal');
if ('IntersectionObserver' in window && revealItems.length) {
    const observer = new IntersectionObserver(
        (entries) => {
            entries.forEach((entry) => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    observer.unobserve(entry.target);
                }
            });
        },
        { threshold: 0.15 }
    );

    revealItems.forEach((item) => observer.observe(item));
} else {
    revealItems.forEach((item) => item.classList.add('visible'));
}

const milestoneSelect = document.getElementById('milestoneSelect');
const milestoneName = document.getElementById('milestoneName');
const milestoneDesc = document.getElementById('milestoneDesc');
const milestoneDate = document.getElementById('milestoneDate');
const milestoneMarks = document.getElementById('milestoneMarks');
const milestoneStatus = document.getElementById('milestoneStatus');
const milestoneDateWrap = document.getElementById('milestoneDateWrap');
const milestoneProgress = document.getElementById('milestoneProgress');

const milestoneData = {
    proposal: {
        title: 'Project Proposal',
        description: 'Initial proposal covering problem statement, objectives, methodology, and feasibility. Presented to project supervisors and panel.',
        date: '',
        marks: '6%',
        status: 'Completed',
        progress: 100,
        statusClass: 'is-completed'
    },
    pp1: {
        title: 'Progress Presentation 1',
        description: 'First progress review covering system architecture, initial implementation, and literature survey findings.',
        date: '07 January 2025',
        marks: '15%',
        status: 'Completed',
        progress: 100,
        statusClass: 'is-completed'
    },
    pp2: {
        title: 'Progress Presentation 2',
        description: 'Current stage - demonstrating 80% complete platform with all four components integrated. System architecture and API endpoints presented.',
        date: '10 March 2026',
        marks: '18%',
        status: 'Completed',
        progress: 100,
        statusClass: 'is-completed'
    },
    viva: {
        title: 'Final Viva',
        description: 'Final system demonstration with complete documentation, research paper, and fully deployed platform evaluation.',
        date: '05 May 2026',
        marks: '10%',
        status: 'Upcoming',
        progress: 20,
        statusClass: 'is-upcoming'
    }
};

if (milestoneSelect) {
    const renderMilestone = (selected) => {
        if (!selected) {
            milestoneName.textContent = 'Select an assessment';
            milestoneDesc.textContent = 'Choose a milestone from the dropdown to view details.';
            milestoneDate.textContent = '';
            milestoneMarks.textContent = '';
            milestoneStatus.textContent = '';
            if (milestoneDateWrap) {
                milestoneDateWrap.hidden = false;
            }
            if (milestoneProgress) {
                milestoneProgress.style.width = '0%';
            }
            return;
        }

        milestoneName.textContent = selected.title;
        milestoneDesc.textContent = selected.description;
        milestoneDate.textContent = selected.date;
        milestoneMarks.textContent = selected.marks;
        milestoneStatus.textContent = selected.status;

        if (milestoneDateWrap) {
            milestoneDateWrap.hidden = !selected.date;
        }

        if (milestoneStatus) {
            milestoneStatus.classList.remove('is-completed', 'is-upcoming');
            if (selected.statusClass) {
                milestoneStatus.classList.add(selected.statusClass);
            }
        }

        if (milestoneProgress) {
            milestoneProgress.style.width = `${selected.progress}%`;
        }
    };

    milestoneSelect.addEventListener('change', (event) => {
        renderMilestone(milestoneData[event.target.value]);
    });

    renderMilestone(milestoneData[milestoneSelect.value]);
}

const slideEmbeds = document.querySelectorAll('.slide-embed[data-src]');
if (slideEmbeds.length) {
    slideEmbeds.forEach((frame) => {
        const source = frame.getAttribute('data-src');
        if (!source) {
            return;
        }

        // Use Office online viewer so previews are rendered from the actual slide links.
        frame.src = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(source)}`;
    });
}

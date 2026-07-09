/**
 * Fintech Payment OpenAPI Docs - Main JavaScript
 */

(function() {
  'use strict';

  // Mobile menu toggle
  var menuToggle = document.querySelector('.menu-toggle');
  var sidebar = document.querySelector('.sidebar');
  function syncSidebarState() {
    if (sidebar.classList.contains('open')) {
      document.body.classList.add('sidebar-open');
    } else {
      document.body.classList.remove('sidebar-open');
    }
  }
  if (menuToggle && sidebar) {
    menuToggle.addEventListener('click', function(e) {
      e.stopPropagation();
      sidebar.classList.toggle('open');
      syncSidebarState();
    });
    // Close sidebar when clicking the backdrop or outside on mobile
    document.addEventListener('click', function(e) {
      if (window.innerWidth <= 900 && sidebar.classList.contains('open')) {
        if (!sidebar.contains(e.target) && !menuToggle.contains(e.target)) {
          sidebar.classList.remove('open');
          syncSidebarState();
        }
      }
    });
    // Close sidebar after tapping a link on mobile (better UX)
    sidebar.querySelectorAll('a').forEach(function(link) {
      link.addEventListener('click', function() {
        if (window.innerWidth <= 900) {
          sidebar.classList.remove('open');
          syncSidebarState();
        }
      });
    });
  }

  // Sidebar group collapse/expand - only toggle on the icon, not the whole title
  document.querySelectorAll('.sidebar-group-title').forEach(function(title) {
    title.addEventListener('click', function(e) {
      // Only toggle if clicking the toggle-icon or the title text (not child links)
      if (e.target.closest('.sidebar-links')) return;
      var group = this.closest('.sidebar-group');
      if (group) {
        group.classList.toggle('collapsed');
      }
    });
  });

  // Ensure all groups are expanded by default on first visit
  document.querySelectorAll('.sidebar-group').forEach(function(group) {
    group.classList.remove('collapsed');
  });

  // TOC active section highlighting on scroll
  var tocNav = document.querySelector('.toc-nav');
  if (tocNav) {
    var headings = document.querySelectorAll('.main-content h2, .main-content h3, .main-content h4');
    var tocLinks = tocNav.querySelectorAll('a');

    function highlightTOC() {
      var current = '';
      var scrollPos = window.scrollY + 100;

      headings.forEach(function(heading) {
        if (heading.offsetTop <= scrollPos) {
          current = heading.id;
        }
      });

      tocLinks.forEach(function(link) {
        link.classList.remove('active');
        if (link.getAttribute('href') === '#' + current) {
          link.classList.add('active');
        }
      });
    }

    window.addEventListener('scroll', highlightTOC);
    highlightTOC();
  }

  // Copy code blocks
  document.querySelectorAll('pre').forEach(function(pre) {
    // Skip if already has a copy button
    if (pre.querySelector('.copy-btn')) return;

    var btn = document.createElement('button');
    btn.className = 'copy-btn';
    btn.textContent = (document.documentElement.lang === 'en') ? 'Copy' : '複製';
    btn.type = 'button';
    pre.appendChild(btn);

    pre.addEventListener('mouseenter', function() { btn.style.opacity = '1'; });
    pre.addEventListener('mouseleave', function() { btn.style.opacity = '0'; });

    btn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      var code = pre.querySelector('code');
      if (code) {
        navigator.clipboard.writeText(code.textContent).then(function() {
          btn.textContent = (document.documentElement.lang === 'en') ? 'Copied' : '已複製';
          setTimeout(function() { btn.textContent = (document.documentElement.lang === 'en') ? 'Copy' : '複製'; }, 2000);
        });
      }
    });
  });

  // Smooth scroll for anchor links (TOC links only)
  document.querySelectorAll('.toc-nav a[href^="#"]').forEach(function(anchor) {
    anchor.addEventListener('click', function(e) {
      var targetId = this.getAttribute('href');
      if (targetId === '#') return;
      var target = document.querySelector(targetId);
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: 'smooth' });
        history.pushState(null, null, targetId);
      }
    });
  });

  // Highlight current page in sidebar
  var currentPath = window.location.pathname;
  document.querySelectorAll('.sidebar-links a').forEach(function(link) {
    var href = link.getAttribute('href');
    if (!href) return;
    // Normalize: remove all ../ and ./ prefixes
    var normalized = href.replace(/^(?:\.\.\/)+/, '').replace(/^\.\//, '');
    // Check if current path ends with this normalized path
    if (currentPath.endsWith(normalized) || currentPath.endsWith('/' + normalized)) {
      link.classList.add('active');
      // Expand parent group
      var group = link.closest('.sidebar-group');
      if (group) {
        group.classList.remove('collapsed');
      }
    } else {
      link.classList.remove('active');
    }
  });

})();

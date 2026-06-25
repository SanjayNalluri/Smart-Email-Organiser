import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import Sidebar from '../components/Sidebar';
import CategoryCard from '../components/CategoryCard';
import EmailList from '../components/EmailList';
import EmailDetail from '../components/EmailDetail';
import ChatBot from '../components/ChatBot';
import EventCalendar from '../components/EventCalendar';
import { getCurrentUser, getEmails, logout } from '../services/api';

const CATEGORIES = ['Events', 'Academics', 'Hackathons', 'Personal', 'Spam'];

export default function Dashboard() {
    const navigate = useNavigate();

    const [user, setUser] = useState(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);
    const [emailData, setEmailData] = useState({ emails: [], grouped: {} });
    const [selectedCategory, setSelectedCategory] = useState(null);
    const [selectedEmail, setSelectedEmail] = useState(null);
    const [searchQuery, setSearchQuery] = useState('');

    // Handle re-categorization — updates local state when user corrects a category
    const handleRecategorize = (emailId, newCategory) => {
        setEmailData(prev => {
            const updatedEmails = prev.emails.map(e =>
                e.id === emailId ? { ...e, category: newCategory } : e
            );
            const grouped = updatedEmails.reduce((acc, email) => {
                if (!acc[email.category]) acc[email.category] = [];
                acc[email.category].push(email);
                return acc;
            }, {});
            return { emails: updatedEmails, grouped };
        });
        // Update selected email too
        setSelectedEmail(prev =>
            prev && prev.id === emailId ? { ...prev, category: newCategory } : prev
        );
    };

    useEffect(() => {
        const checkAuth = async () => {
            try {
                const { user, isAuthenticated } = await getCurrentUser();
                if (!isAuthenticated) {
                    navigate('/login');
                    return;
                }
                setUser(user);
                const data = await getEmails();
                if (data && data.error) {
                    setError(data.error);
                } else {
                    setEmailData(data || { emails: [], grouped: {} });
                }
            } catch (err) {
                console.error('Auth check failed:', err);
                setError('Failed to connect to the server.');
            } finally {
                setLoading(false);
            }
        };
        checkAuth();
    }, [navigate]);

    const handleBackToCategories = () => {
        setSelectedCategory(null);
        setSelectedEmail(null);
        setSearchQuery('');
    };

    // Filter emails based on search query — searches subject, snippet, from, and body
    const filterEmails = (emails) => {
        if (!searchQuery.trim()) return emails;
        const query = searchQuery.toLowerCase();
        return emails.filter(email => {
            const plainBody = email.body
                ? email.body.replace(/<[^>]*>/g, ' ').toLowerCase()
                : '';
            return (
                email.subject?.toLowerCase().includes(query) ||
                email.snippet?.toLowerCase().includes(query) ||
                email.from?.toLowerCase().includes(query) ||
                plainBody.includes(query)
            );
        });
    };

    // Get search results across all emails
    const getSearchResults = () => {
        if (!searchQuery.trim()) return null;
        return filterEmails(emailData.emails);
    };



    if (loading) {
        return (
            <div className="app-layout">
                <Sidebar user={null} />
                <main className="main-content">
                    <div className="loading"><div className="loading__spinner" /></div>
                </main>
            </div>
        );
    }

    if (error) {
        return (
            <div className="app-layout">
                <Sidebar user={user} />
                <main className="main-content">
                    <div className="error-container">
                        <div className="error-card">
                            <div className="error-card__icon">
                                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#ea4335" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                            </div>
                            <h2 className="error-card__title">Failed to load emails</h2>
                            <p className="error-card__message">
                                {error === 'Failed to fetch emails'
                                    ? 'The Gmail API returned a permission error. This usually means the Google account lacks access or you did not check the "Read your emails" permission on the Google Sign-In screen.'
                                    : error}
                            </p>
                            <div className="error-card__actions">
                                <button onClick={() => window.location.reload()} className="error-card__btn error-card__btn--primary">
                                    Retry Connection
                                </button>
                                <button onClick={async () => {
                                    await logout();
                                    navigate('/login');
                                }} className="error-card__btn error-card__btn--danger">
                                    Sign Out / Switch Account
                                </button>
                            </div>
                        </div>
                    </div>
                </main>
            </div>
        );
    }

    const searchResults = getSearchResults();
    const displayEmails = searchResults || (selectedCategory ? emailData.grouped[selectedCategory] || [] : emailData.emails);

    return (
        <div className="app-layout">
            <Sidebar
                user={user}
                categories={CATEGORIES}
                grouped={emailData.grouped}
                onSelectCategory={(cat) => {
                    setSelectedCategory(cat);
                    setSelectedEmail(null);
                    setSearchQuery('');
                }}
                selectedCategory={selectedCategory}
            />
            <main className="main-content">
                <header className="dashboard-header">
                    <h1 className="dashboard-header__title">
                        {selectedCategory ? (
                            <>
                                <button onClick={handleBackToCategories} className="back-btn">←</button>
                                {selectedCategory}
                            </>
                        ) : 'Dashboard'}
                    </h1>
                    <p className="dashboard-header__subtitle">
                        {searchQuery
                            ? `Found ${searchResults?.length || 0} emails matching "${searchQuery}"`
                            : selectedCategory
                                ? `${displayEmails.length} emails in this category`
                                : `You have ${emailData.emails.length} emails organized across ${Object.keys(emailData.grouped).length} categories.`}
                    </p>
                </header>



                {/* Search Bar */}
                <div className="search-container">
                    <div className="search-bar">
                        <svg className="search-icon" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <circle cx="11" cy="11" r="8"></circle>
                            <path d="m21 21-4.35-4.35"></path>
                        </svg>
                        <input
                            type="text"
                            className="search-input"
                            placeholder="Search emails by subject, content, or sender..."
                            value={searchQuery}
                            onChange={(e) => setSearchQuery(e.target.value)}
                        />
                        {searchQuery && (
                            <button
                                className="search-clear"
                                onClick={() => setSearchQuery('')}
                                aria-label="Clear search"
                            >
                                ×
                            </button>
                        )}
                    </div>
                </div>

                {/* Show search results */}
                {searchQuery && searchResults && (
                    <div className="email-section">
                        <EmailList
                            emails={searchResults}
                            selectedEmail={selectedEmail}
                            onSelectEmail={setSelectedEmail}
                            title={`Search Results for "${searchQuery}"`}
                        />
                        <EmailDetail email={selectedEmail} onRecategorize={handleRecategorize} />
                    </div>
                )}

                {/* Show categories when not searching */}
                {!searchQuery && !selectedCategory && (
                    <div className="bento-grid">
                        {CATEGORIES.map((category) => (
                            <CategoryCard
                                key={category}
                                category={category}
                                emails={emailData.grouped[category] || []}
                                onClick={() => setSelectedCategory(category)}
                            />
                        ))}
                    </div>
                )}

                {/* Calendar Widget */}
                {!searchQuery && !selectedCategory && emailData.emails.length > 0 && (
                    <EventCalendar 
                        emails={emailData.emails} 
                        onEventClick={(event) => {
                            setSelectedCategory(event.category);
                            setSelectedEmail(event);
                        }}
                    />
                )}

                {!searchQuery && selectedCategory && (
                    <div className="email-section">
                        <EmailList
                            emails={displayEmails}
                            selectedEmail={selectedEmail}
                            onSelectEmail={setSelectedEmail}
                            title={selectedCategory}
                        />
                        <EmailDetail email={selectedEmail} onRecategorize={handleRecategorize} />
                    </div>
                )}

                {!searchQuery && !selectedCategory && emailData.emails.length > 0 && (
                    <>
                        <h2 style={{ marginBottom: '1.5rem', marginTop: '0.5rem' }}>Recent Emails</h2>
                        <div className="email-section">
                            <EmailList
                                emails={emailData.emails.slice(0, 10)}
                                selectedEmail={selectedEmail}
                                onSelectEmail={setSelectedEmail}
                                title="All Categories"
                            />
                            <EmailDetail email={selectedEmail} onRecategorize={handleRecategorize} />
                        </div>
                    </>
                )}
            </main>
            <ChatBot
                emails={emailData.emails}
                onNavigateCategory={(category) => {
                    // Navigate to a category
                    const matchedCategory = CATEGORIES.find(c => c.toLowerCase() === category.toLowerCase());
                    if (matchedCategory) {
                        setSelectedCategory(matchedCategory);
                        setSelectedEmail(null);
                        setSearchQuery('');
                    }
                }}
                onSearch={(query) => {
                    // Fill search bar and show results
                    setSearchQuery(query);
                    setSelectedCategory(null);
                    setSelectedEmail(null);
                }}
                onSelectEmail={(emailSubject) => {
                    // Find and select an email by subject match
                    const subject = emailSubject.toLowerCase();
                    const found = emailData.emails.find(e =>
                        e.subject?.toLowerCase().includes(subject)
                    );
                    if (found) {
                        setSelectedCategory(found.category);
                        setSelectedEmail(found);
                        setSearchQuery('');
                    }
                }}
                onGoDashboard={() => {
                    // Return to main dashboard view
                    setSelectedCategory(null);
                    setSelectedEmail(null);
                    setSearchQuery('');
                }}
                onRecategorize={handleRecategorize}
            />
        </div>
    );
}
